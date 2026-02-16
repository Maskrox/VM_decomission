# =============================================================================
# DecomWorker.ps1
#
# This script is designed to be executed as a PowerShell Job by the
# VM_LIFECYCLE_MANAGER.ps1 GUI application. It performs the decommissioning
# tasks for a single target machine in parallel.
#
# =============================================================================

# --- LOAD MODULES & SNAPINS ---
# Import core PowerShell modules required for infrastructure management.
Import-Module ActiveDirectory, VirtualMachineManager -ErrorAction SilentlyContinue

# Citrix Modules/Snapins - Depending on XenApp/XenDesktop version, one or more may be needed.
# This section attempts to load Citrix PowerShell components, prioritizing snapins for older
# environments and falling back to modules for newer SDKs.
try {
    Add-PSSnapin Citrix* -ErrorAction SilentlyContinue
} catch {
    # Attempt to import common Citrix modules for modern SDKs
    Import-Module Broker -ErrorAction SilentlyContinue
    Import-Module Configuration -ErrorAction SilentlyContinue
    Import-Module Host -ErrorAction SilentlyContinue
    Import-Module DelegatedAdmin -ErrorAction SilentlyContinue
}

# --- Simplified Log-Msg for Worker ---
# This function is a simplified logging mechanism for the worker script.
# It uses Write-Host to output messages, which will be captured by the PowerShell job's output stream.
function Log-Msg {
    param (
        [string]$Msg,
        [string]$ColorName="White" # Color is not functional for Write-Host in jobs but kept for consistency
    )
    # Prefix worker messages to distinguish them in job output
    Write-Host "[WORKER] [$(Get-Date -Format 'HH:mm:ss')] $Msg"
}

# --- Stop Target System Function ---
# Attempts to shut down a target machine, either physical or virtual (VMM).
# This function is intended to be run in a PowerShell Job.
function Stop-TargetSystem {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Target,
        [Parameter(Mandatory=$true)]
        [string]$Prefix,
        [Parameter(Mandatory=$true)]
        [bool]$IsPhysical,
        [Parameter(Mandatory=$true)]
        [hashtable]$VmmTargetHost # Hashtable for VMM host mapping (Target => VMM Server)
    )
    $success = $false
    $errorMessage = $null

    if ($IsPhysical) {
        try { 
            Log-Msg "${Prefix}[$Target] Attempting to shut down physical machine." # Use Log-Msg for worker output
            Stop-Computer -ComputerName $Target -Force -ErrorAction Stop
            Log-Msg "${Prefix}[$Target] Shutdown signal sent to physical machine."
            $success = $true
        } catch { 
            $errorMessage = "Error stopping physical machine via WMI: $($_.Exception.Message)"
            Log-Msg "${Prefix}[$Target] $errorMessage"
        }
    } else {
        $VMM = $VmmTargetHost[$Target] # Get VMM host for the target VM.
        if ($VMM) {
            try { 
                Log-Msg "${Prefix}[$Target] Attempting to shut down VMM VM on $VMM."
                Stop-SCVirtualMachine -VMMServer $VMM -VM $Target -Confirm:$false -ErrorAction Stop
                Log-Msg "${Prefix}[$Target] VMM VM shutdown complete."
                $success = $true
            } catch { 
                $errorMessage = "Error stopping VMM VM: $($_.Exception.Message)"
                Log-Msg "${Prefix}[$Target] $errorMessage"
            }
        } else { 
            $errorMessage = "Skip: Not mapped to VMM, cannot stop virtual machine."
            Log-Msg "${Prefix}[$Target] $errorMessage"
        }
    }
    return [PSCustomObject]@{
        Target = $Target
        Phase = "Stop"
        Success = $success
        Message = if ($success) { "Stop operation for $Target completed successfully." } else { $errorMessage }
    }
}

# --- Clean Active Directory Function ---
# Attempts to find and delete the computer object from Active Directory.
# This function is intended to be run in a PowerShell Job.
function Clean-TargetAD {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Target,
        [Parameter(Mandatory=$true)]
        [string]$Prefix,
        [Parameter(Mandatory=$true)]
        [string]$TargetOU # Active Directory OU for computer object lookup/deletion
    )
    $success = $false
    $errorMessage = $null

    try {
        Log-Msg "${Prefix}[$Target] Attempting to clean AD entry."
        $Searcher = [ADSISearcher]""
        $Searcher.SearchRoot = [ADSI]"LDAP://$TargetOU"
        $Searcher.Filter = "(&(objectClass=computer)(name=$Target))"
        $Result = $Searcher.FindOne()
        if ($Result) {
            $DN = $Result.Properties.distinguishedname[0]
            $ADSIObj = [ADSI]"LDAP://$DN"
            $ADSIObj.DeleteTree()
            Log-Msg "${Prefix}[$Target] AD: Computer object '$DN' purged."
            $success = $true
        } else {
            $errorMessage = "AD: Computer object not found in target OU '$TargetOU'. Skipping."
            Log-Msg "${Prefix}[$Target] $errorMessage"
            $success = $true # Not found is not an error for cleanup, consider it successful in this context
        }
    } catch { 
        $errorMessage = "AD: Error during deletion: $($_.Exception.Message)"
        Log-Msg "${Prefix}[$Target] $errorMessage"
    }
    return [PSCustomObject]@{
        Target = $Target
        Phase = "Clean-AD"
        Success = $success
        Message = if ($success) { "AD cleanup for $Target completed successfully." } else { $errorMessage }
    }
}

# --- Clean SCCM Function ---
# Attempts to remove the device from SCCM.
# This function is intended to be run in a PowerShell Job.
function Clean-TargetSCCM {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Target,
        [Parameter(Mandatory=$true)]
        [string]$Prefix,
        [Parameter(Mandatory=$true)]
        [string]$SccmSiteCode,  # SCCM Site Code
        [Parameter(Mandatory=$true)]
        [string]$SccmProvider   # SCCM SMS Provider server
    )
    $success = $false
    $errorMessage = $null

    try {
        Log-Msg "${Prefix}[$Target] Attempting to clean SCCM entry."
        # SCCM Module requires a PSDrive setup. We manage this locally for the job.
        if (!(Test-Path "$($SccmSiteCode):")) { 
            # Create a temporary PSDrive for SCCM if it doesn't exist
            New-PSDrive -Name $SccmSiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root $SccmProvider -ErrorAction Stop | Out-Null 
            Log-Msg "${Prefix}[$Target] SCCM: Created PSDrive for site '$SccmSiteCode'."
        }
        $CurrentLocation = (Get-Location).Path # Store current location
        Set-Location "$($SccmSiteCode):" # Change to SCCM PSDrive
        Remove-CMDevice -DeviceName $Target -Force -Confirm:$false -ErrorAction Stop
        Set-Location $CurrentLocation # Revert to original location
        Log-Msg "${Prefix}[$Target] SCCM: Device removed."
        $success = $true
    } catch { 
        $errorMessage = "SCCM: Error during removal or not found: $($_.Exception.Message)"
        Log-Msg "${Prefix}[$Target] $errorMessage" 
        # Ensure we exit the SCCM drive if an error occurs and location was changed.
        if ((Get-Location).Path -like "$($SccmSiteCode):*") {
            Set-Location "C:" 
        }
    }
    return [PSCustomObject]@{
        Target = $Target
        Phase = "Clean-SCCM"
        Success = $success
        Message = if ($success) { "SCCM cleanup for $Target completed successfully." } else { $errorMessage }
    }
}

# --- Clean Citrix Function ---
# Performs a comprehensive cleanup of a machine from Citrix Virtual Apps and Desktops.
# This function is intended to be run in a PowerShell Job.
function Clean-TargetCitrix {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Target,
        [Parameter(Mandatory=$true)]
        [string]$Prefix,
        [Parameter(Mandatory=$true)]
        [bool]$IncludeCitrixCleanup,
        [Parameter(Mandatory=$true)]
        [bool]$IsPhysical
        # No need for CitrixDDC if cmdlets are properly loaded in worker script.
    )
    $success = $false
    $errorMessage = $null

    if (!$IncludeCitrixCleanup) {
        $errorMessage = "Citrix: Cleanup skipped as per user selection."
        Log-Msg "${Prefix}[$Target] $errorMessage"
        return [PSCustomObject]@{
            Target = $Target
            Phase = "Clean-Citrix"
            Success = $true # Skipped is not a failure
            Message = $errorMessage
        }
    }

    if ($IsPhysical) {
        $errorMessage = "Citrix: Cleanup skipped (Physical Mode)."
        Log-Msg "${Prefix}[$Target] $errorMessage"
        return [PSCustomObject]@{
            Target = $Target
            Phase = "Clean-Citrix"
            Success = $true # Skipped is not a failure
            Message = $errorMessage
        }
    }

    Log-Msg "${Prefix}[$Target] Citrix: Starting cleanup sequence."
    try {
        # 1. Get Citrix Machine Object: Retrieve the machine object from Citrix.
        $CitrixMachine = Get-BrokerMachine -MachineName $Target -ErrorAction SilentlyContinue
        
        if ($CitrixMachine) {
            Log-Msg "${Prefix}[$Target] Citrix: Machine found (UUID: $($CitrixMachine.MachineUid))."

            # 2. Put into Maintenance Mode: Set the machine to maintenance mode if it's not powered off or already in maintenance.
            if ($CitrixMachine.PowerState -ne "PoweredOff" -and $CitrixMachine.InMaintenanceMode -eq $false) {
                Log-Msg "${Prefix}[$Target] Citrix: Setting machine to maintenance mode."
                Set-BrokerMachine -InputObject $CitrixMachine -InMaintenanceMode $true -ErrorAction Stop
                Log-Msg "${Prefix}[$Target] Citrix: Machine now in maintenance mode."
            } elseif ($CitrixMachine.InMaintenanceMode -eq $true) {
                Log-Msg "${Prefix}[$Target] Citrix: Machine already in maintenance mode."
            } else {
                Log-Msg "${Prefix}[$Target] Citrix: Machine is powered off, skipping maintenance mode setting."
            }

            # 3. Remove from Delivery Groups: Remove the machine from all associated Delivery Groups.
            try {
                $DeliveryGroups = Get-BrokerDeliveryGroup -MachineFilter "MachineName -eq '$Target'" -ErrorAction SilentlyContinue
                if ($DeliveryGroups) {
                    foreach ($DG in $DeliveryGroups) {
                        Log-Msg "${Prefix}[$Target] Citrix: Removing from Delivery Group $($DG.Name)."
                        Remove-BrokerMachineFromDeliveryGroup -InputObject $CitrixMachine -DeliveryGroup $DG -ErrorAction Stop
                        Log-Msg "${Prefix}[$Target] Citrix: Removed from Delivery Group $($DG.Name)."
                    }
                } else {
                    Log-Msg "${Prefix}[$Target] Citrix: Not found in any Delivery Group."
                }
            } catch {
                $errorMessage = "Citrix: Error removing from Delivery Group: $($_.Exception.Message)"
                Log-Msg "${Prefix}[$Target] $errorMessage"
            }

            # 4. Remove from Machine Catalog: Remove the machine from its Machine Catalog.
            try {
                $MachineCatalog = Get-BrokerMachineCatalog -MachineFilter "MachineName -eq '$Target'" -ErrorAction SilentlyContinue
                if ($MachineCatalog) {
                    Log-Msg "${Prefix}[$Target] Citrix: Removing from Machine Catalog $($MachineCatalog.Name)."
                    Remove-BrokerMachineFromMachineCatalog -InputObject $CitrixMachine -MachineCatalog $MachineCatalog -ErrorAction Stop
                    Log-Msg "${Prefix}[$Target] Citrix: Removed from Machine Catalog $($MachineCatalog.Name)."
                } else {
                    Log-Msg "${Prefix}[$Target] Citrix: Not found in any Machine Catalog."
                }
            } catch {
                $errorMessage = "Citrix: Error removing from Machine Catalog: $($_.Exception.Message)"
                Log-Msg "${Prefix}[$Target] $errorMessage"
            }

            # 5. Delete Machine from Citrix: Delete the machine object from the Citrix site.
            try {
                Log-Msg "${Prefix}[$Target] Citrix: Deleting machine object from Citrix."
                Remove-BrokerMachine -InputObject $CitrixMachine -ErrorAction Stop
                Log-Msg "${Prefix}[$Target] Citrix: Machine object deleted successfully."
                $success = $true
            } catch {
                $errorMessage = "Citrix: Error deleting machine object: $($_.Exception.Message)"
                Log-Msg "${Prefix}[$Target] $errorMessage"
            }

            Log-Msg "${Prefix}[$Target] Citrix: All cleanup steps complete."
        } else {
            $errorMessage = "Citrix: Machine object not found in Citrix. Skipping Citrix cleanup."
            Log-Msg "${Prefix}[$Target] $errorMessage"
            $success = $true # Not found is not an error for cleanup, consider it successful in this context
        }
    } catch {
        $errorMessage = "Citrix: General error during cleanup: $($_.Exception.Message)"
        Log-Msg "${Prefix}[$Target] $errorMessage"
    }
    return [PSCustomObject]@{
        Target = $Target
        Phase = "Clean-Citrix"
        Success = $success
        Message = if ($success) { "Citrix cleanup for $Target completed successfully." } else { $errorMessage }
    }
}

# --- Clean DNS Function ---
# Attempts to remove the DNS 'A' record for the target machine.
# This function is intended to be run in a PowerShell Job.
function Clean-TargetDNS {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Target,
        [Parameter(Mandatory=$true)]
        [string]$Prefix,
        [Parameter(Mandatory=$true)]
        [string]$DC, # Domain Controller for DNS operations
        [Parameter(Mandatory=$true)]
        [string]$ForwardZone # DNS Forward Lookup Zone
    )
    $success = $false
    $errorMessage = $null

    try {
        Log-Msg "${Prefix}[$Target] Attempting to clean DNS entry."
        Remove-DnsServerResourceRecord -ComputerName $DC -ZoneName $ForwardZone -Name $Target -RRType "A" -Force -ErrorAction Stop
        Log-Msg "${Prefix}[$Target] DNS: 'A' record removed."
        $success = $true
    } catch { 
        $errorMessage = "DNS: Error during removal: $($_.Exception.Message)"
        Log-Msg "${Prefix}[$Target] $errorMessage"
    }
    return [PSCustomObject]@{
        Target = $Target
        Phase = "Clean-DNS"
        Success = $success
        Message = if ($success) { "DNS cleanup for $Target completed successfully." } else { $errorMessage }
    }
}

# --- Delete Target VM Function ---
# Attempts to delete a virtual machine from the VMM hypervisor.
# This function is intended to be run in a PowerShell Job.
function Delete-TargetVM {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Target,
        [Parameter(Mandatory=$true)]
        [string]$Prefix,
        [Parameter(Mandatory=$true)]
        [bool]$IsPhysical,
        [Parameter(Mandatory=$true)]
        [hashtable]$VmmTargetHost # Hashtable for VMM host mapping (Target => VMM Server)
    )
    $success = $false
    $errorMessage = $null

    # Double safety check for physical mode, though button should be disabled
    if ($IsPhysical) { 
        $errorMessage = "ABORTED: Physical hardware protection. Cannot delete from hypervisor."
        Log-Msg "${Prefix}[$Target] $errorMessage"
        return [PSCustomObject]@{
            Target = $Target
            Phase = "Delete-VM"
            Success = $false
            Message = $errorMessage
        }
    }

    $VMM = $VmmTargetHost[$Target]
    if ($VMM) {
        try { 
            Log-Msg "${Prefix}[$Target] Deleting VM from VMM on host $VMM..."
            Remove-SCVirtualMachine -VMMServer $VMM -VM $Target -Force -Confirm:$false -ErrorAction Stop
            Log-Msg "${Prefix}[$Target] VMM: VM DELETED PERMANENTLY."
            $success = $true
        } catch { 
            $errorMessage = "VMM: Error deleting VM: $($_.Exception.Message)"
            Log-Msg "${Prefix}[$Target] $errorMessage"
        }
    } else {
        $errorMessage = "Skip: VMM Host unknown for '$Target'. Cannot delete VM."
        Log-Msg "${Prefix}[$Target] $errorMessage"
    }
    return [PSCustomObject]@{
        Target = $Target
        Phase = "Delete-VM"
        Success = $success
        Message = if ($success) { "VM deletion for $Target completed successfully." } else { $errorMessage }
    }
}
