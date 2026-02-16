# =============================================================================
# VM LIFECYCLE MANAGER - ENTERPRISE EDITION (V6.1)
# =============================================================================

# --- 1. HIGH DPI AWARENESS ---
try {
    $code = '[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();'
    $Win32 = Add-Type -MemberDefinition $code -Name "Win32" -Namespace Win32 -PassThru
    $Win32::SetProcessDPIAware() | Out-Null
} catch { }

# --- LOAD ASSEMBLIES & MODULES ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
Import-Module ActiveDirectory, VirtualMachineManager, ConfigurationManager -ErrorAction SilentlyContinue

# --- GLOBAL VARIABLES (INITIALIZE EMPTY) ---
$Script:RunCreds = $null 
$Script:VmmTargetHost = @{}
$Script:IsPhysical = $false

# Infrastructure Variables (Script Scope for dynamic updates)
$Script:VmmServers    = @()
$Script:DC            = ""
$Script:SccmSiteCode  = ""
$Script:SccmProvider  = ""
$Script:ForwardZone   = ""
$Script:TargetOU      = ""
$Script:CitrixDDC     = ""

# --- THEME DEFINITION (HARDCODED DEFAULT) ---
$Theme = @{
    BgBase      = [System.Drawing.Color]::FromArgb(32, 32, 32)
    BgControl   = [System.Drawing.Color]::FromArgb(45, 45, 48)
    TextMain    = [System.Drawing.Color]::FromArgb(240, 240, 240)
    TextDim     = [System.Drawing.Color]::FromArgb(160, 160, 160)
    AccentBlue  = [System.Drawing.Color]::FromArgb(0, 120, 215)
    AccentGreen = [System.Drawing.Color]::FromArgb(16, 124, 16)
    AccentRed   = [System.Drawing.Color]::FromArgb(232, 17, 35)
    AccentWarn  = [System.Drawing.Color]::FromArgb(255, 140, 0)
    Border      = [System.Drawing.Color]::FromArgb(80, 80, 80)
}

$FontTitle = New-Object System.Drawing.Font("Segoe UI Semibold", 18)
$FontHead  = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$FontText  = New-Object System.Drawing.Font("Segoe UI", 10)
$FontSmall = New-Object System.Drawing.Font("Segoe UI", 8)
$FontMono  = New-Object System.Drawing.Font("Consolas", 9)

# --- UI SETUP ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Server Decommissioning Manager - Enterprise Master"; 
$form.Size = New-Object System.Drawing.Size(950, 950) 
$form.MinimumSize = New-Object System.Drawing.Size(950, 800)
$form.BackColor = $Theme.BgBase
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true

# Main Panel with Scroll Support
$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Dock = "Fill"
$mainPanel.AutoScroll = $true
$mainPanel.AutoScrollMinSize = New-Object System.Drawing.Size(900, 980) 
$form.Controls.Add($mainPanel)

# --- UI HELPER FUNCTIONS ---
function New-StyledButton {
    param ($Text, $Color, $X, $Y, $W, $H, $Parent, $Anchor)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, $Y)
    $btn.Size = New-Object System.Drawing.Size($W, $H)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    if ($Color) { $btn.BackColor = $Color } else { $btn.BackColor = $Theme.BgControl }
    $btn.ForeColor = $Theme.TextMain
    $btn.Font = $FontHead
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    if ($Anchor) { $btn.Anchor = $Anchor }
    $Parent.Controls.Add($btn)
    return $btn
}

function New-StyledGroup {
    param ($Text, $X, $Y, $W, $H, $Parent)
    $grp = New-Object System.Windows.Forms.GroupBox
    $grp.Text = $Text
    $grp.Location = New-Object System.Drawing.Point($X, $Y)
    $grp.Size = New-Object System.Drawing.Size($W, $H)
    $grp.ForeColor = $Theme.AccentBlue
    $grp.Font = $FontHead
    $grp.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $Parent.Controls.Add($grp)
    return $grp
}

function New-DescLabel {
    param ($Text, $X, $Y, $W, $Parent)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.Location = New-Object System.Drawing.Point($X, $Y)
    $lbl.Size = New-Object System.Drawing.Size($W, 20)
    $lbl.ForeColor = $Theme.TextDim
    $lbl.Font = $FontSmall
    $lbl.TextAlign = "TopCenter"
    $Parent.Controls.Add($lbl)
}

# --- GUI CONSTRUCTION ---

# Header Title
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "VM LIFECYCLE MANAGER"
$lblTitle.Font = $FontTitle
$lblTitle.ForeColor = $Theme.AccentBlue
$lblTitle.Location = New-Object System.Drawing.Point(20, 20)
$lblTitle.Size = New-Object System.Drawing.Size(400, 40)
$mainPanel.Controls.Add($lblTitle)

# --- BUTTON: LOAD CONFIG ---
$btnLoadConfig = New-StyledButton -Text "LOAD CONFIG" -Color $Theme.BgControl -X 440 -Y 20 -W 180 -H 35 -Parent $mainPanel -Anchor ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$btnLoadConfig.FlatAppearance.BorderColor = $Theme.TextDim
$btnLoadConfig.FlatAppearance.BorderSize = 1
$btnLoadConfig.Font = $FontText

# Login Button
$btnLogin = New-StyledButton -Text "LOGIN CREDENTIALS" -Color $Theme.AccentBlue -X 640 -Y 20 -W 240 -H 35 -Parent $mainPanel -Anchor ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$btnLogin.Font = $FontText

# GROUP 1: DISCOVERY ZONE
$grpDisc = New-StyledGroup -Text " Phase 1: Identification (Batch Support) " -X 20 -Y 70 -W 880 -H 140 -Parent $mainPanel

$lblVM = New-Object System.Windows.Forms.Label
$lblVM.Text = "Target Hostnames (comma separated):" 
$lblVM.Location = New-Object System.Drawing.Point(20, 40); $lblVM.Size = New-Object System.Drawing.Size(300, 25)
$lblVM.ForeColor = $Theme.TextDim; $lblVM.Font = $FontText
$grpDisc.Controls.Add($lblVM)

$txtVM = New-Object System.Windows.Forms.TextBox
$txtVM.Location = New-Object System.Drawing.Point(20, 65); $txtVM.Size = New-Object System.Drawing.Size(350, 28)
$txtVM.Font = New-Object System.Drawing.Font("Segoe UI", 11); $txtVM.BackColor = $Theme.BgControl; $txtVM.ForeColor = $Theme.TextMain; $txtVM.BorderStyle = "FixedSingle"
$grpDisc.Controls.Add($txtVM)

$lblTick = New-Object System.Windows.Forms.Label
$lblTick.Text = "Ticket / Task #:" 
$lblTick.Location = New-Object System.Drawing.Point(380, 40); $lblTick.Size = New-Object System.Drawing.Size(150, 25)
$lblTick.ForeColor = $Theme.TextDim; $lblTick.Font = $FontText
$grpDisc.Controls.Add($lblTick)

$txtTicket = New-Object System.Windows.Forms.TextBox
$txtTicket.Location = New-Object System.Drawing.Point(380, 65); $txtTicket.Size = New-Object System.Drawing.Size(140, 28)
$txtTicket.Font = New-Object System.Drawing.Font("Segoe UI", 11); $txtTicket.BackColor = $Theme.BgControl; $txtTicket.ForeColor = $Theme.AccentWarn; $txtTicket.BorderStyle = "FixedSingle"
$grpDisc.Controls.Add($txtTicket)

$btnCheck = New-StyledButton -Text "1. AUTO-DISCOVER" -Color $Theme.BgControl -X 540 -Y 63 -W 180 -H 32 -Parent $grpDisc
$btnCheck.FlatAppearance.BorderColor = $Theme.AccentBlue; $btnCheck.FlatAppearance.BorderSize = 1

$chkPhys = New-Object System.Windows.Forms.CheckBox
$chkPhys.Text = "Physical Hardware"
$chkPhys.Location = New-Object System.Drawing.Point(20, 100); $chkPhys.Size = New-Object System.Drawing.Size(240, 25)
$chkPhys.ForeColor = $Theme.TextMain; $chkPhys.Font = $FontText
$grpDisc.Controls.Add($chkPhys)

$chkCitrix = New-Object System.Windows.Forms.CheckBox
$chkCitrix.Text = "Include Citrix Cleanup"
$chkCitrix.Location = New-Object System.Drawing.Point(280, 100); $chkCitrix.Size = New-Object System.Drawing.Size(240, 25)
$chkCitrix.ForeColor = $Theme.AccentWarn; $chkCitrix.Font = $FontText
$grpDisc.Controls.Add($chkCitrix)

# GROUP 2: EXECUTION ZONE
$grpAct = New-StyledGroup -Text " Phase 2: Batch Execution Sequence " -X 20 -Y 230 -W 880 -H 150 -Parent $mainPanel

# Buttons
$btnStop = New-StyledButton -Text "2. STOP SYSTEM" -Color $Theme.BgControl -X 30 -Y 40 -W 190 -H 60 -Parent $grpAct
$btnStop.Enabled = $false; $btnStop.FlatAppearance.BorderColor = $Theme.Border; $btnStop.FlatAppearance.BorderSize = 1
New-DescLabel -Text "(Shutdown GuestOS)" -X 30 -Y 105 -W 190 -Parent $grpAct

$btnDecom = New-StyledButton -Text "3. LOGICAL CLEAN" -Color $Theme.BgControl -X 240 -Y 40 -W 190 -H 60 -Parent $grpAct
$btnDecom.Enabled = $false; $btnDecom.FlatAppearance.BorderColor = $Theme.Border; $btnDecom.FlatAppearance.BorderSize = 1
New-DescLabel -Text "(AD/DNS/SCCM/CTX)" -X 240 -Y 105 -W 190 -Parent $grpAct

$btnBackup = New-StyledButton -Text "4. REMOVE BACKUP" -Color $Theme.BgControl -X 450 -Y 40 -W 190 -H 60 -Parent $grpAct
$btnBackup.Enabled = $false; $btnBackup.FlatAppearance.BorderColor = $Theme.Border; $btnBackup.FlatAppearance.BorderSize = 1
New-DescLabel -Text "(Veeam / Commvault)" -X 450 -Y 105 -W 190 -Parent $grpAct

$btnDel = New-StyledButton -Text "5. DELETE VM" -Color $Theme.BgControl -X 660 -Y 40 -W 190 -H 60 -Parent $grpAct
$btnDel.Enabled = $false; $btnDel.FlatAppearance.BorderColor = $Theme.Border; $btnDel.FlatAppearance.BorderSize = 1
New-DescLabel -Text "(Delete from Hypervisor)" -X 660 -Y 105 -W 190 -Parent $grpAct

# LOGS AREA
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Process Activity Log:"
$lblLog.Location = New-Object System.Drawing.Point(25, 400); $lblLog.Size = New-Object System.Drawing.Size(200, 25)
$lblLog.ForeColor = $Theme.TextDim; $lblLog.Font = $FontText
$mainPanel.Controls.Add($lblLog)

$rtbLog = New-Object System.Windows.Forms.RichTextBox
$rtbLog.Location = New-Object System.Drawing.Point(25, 430); $rtbLog.Size = New-Object System.Drawing.Size(875, 400)
$rtbLog.BackColor = "Black"; $rtbLog.ForeColor = $Theme.AccentGreen; $rtbLog.Font = $FontMono; $rtbLog.ReadOnly = $true; $rtbLog.BorderStyle = "None"
$rtbLog.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$mainPanel.Controls.Add($rtbLog)

$btnSaveLog = New-StyledButton -Text "SAVE LOG TO FILE" -Color $Theme.BgControl -X 25 -Y 850 -W 200 -H 35 -Parent $mainPanel -Anchor ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
$btnSaveLog.Font = $FontText; $btnSaveLog.FlatAppearance.BorderColor = $Theme.AccentBlue; $btnSaveLog.FlatAppearance.BorderSize = 1

# =============================================================================
# BUSINESS LOGIC & FUNCTIONS
# =============================================================================

# --- FIXED LOGGING FUNCTION (NO INVOKE ERROR) ---
function Log-Msg {
    param ([string]$Msg, [string]$ColorName="Lime")
    
    # Define the update action
    $Action = {
        $rtbLog.SelectionStart = $rtbLog.TextLength; 
        $rtbLog.SelectionColor = [System.Drawing.Color]::FromName($ColorName)
        $rtbLog.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $Msg`r`n"); 
        $rtbLog.ScrollToCaret()
    }

    # Only Invoke if Handle is created AND Invoke is required
    if ($rtbLog.IsHandleCreated -and $rtbLog.InvokeRequired) {
        $rtbLog.Invoke($Action)
    } else {
        # Run directly (Startup or Main Thread)
        & $Action
    }
}

function Get-TicketPrefix {
    if ($txtTicket.Text.Trim()) { return "[$($txtTicket.Text.Trim())] " }
    return ""
}

function Get-TargetList {
    return $txtVM.Text -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
}

# --- CONFIG LOADING LOGIC (FUNCTION WRAPPED) ---
function Load-InfrastructureConfig ($Path) {
    if (Test-Path $Path) {
        try {
            [xml]$XmlConfig = Get-Content $Path
            $Infra = $XmlConfig.Configuration.Infrastructure
            
            # Update Script Global Variables
            $Script:VmmServers    = @($Infra.VmmServers.Server)
            $Script:DC            = $Infra.DomainController
            $Script:SccmSiteCode  = $Infra.SccmSiteCode
            $Script:SccmProvider  = $Infra.SccmProvider
            $Script:ForwardZone   = $Infra.DnsZone
            $Script:TargetOU      = $Infra.TargetOU
            $Script:CitrixDDC     = $Infra.CitrixController
            
            Log-Msg "--- Configuration Loaded Successfully from: $Path ---" "Cyan"
            Log-Msg "   > Target OU: $Script:TargetOU" "Cyan"
            Log-Msg "   > VMM Servers: $($Script:VmmServers -join ', ')" "Cyan"
            return $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error parsing XML. Check format.", "Config Error", 0, 16)
            Log-Msg "Failed to load config: $($_.Exception.Message)" "Red"
            return $false
        }
    } else {
        return $false
    }
}

# --- BUTTON EVENT HANDLERS ---

# 1. Load Config Button Action
$btnLoadConfig.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "XML Configuration (*.xml)|*.xml|All Files (*.*)|*.*"
    $ofd.Title = "Select Infrastructure Configuration"
    
    if ($ofd.ShowDialog() -eq "OK") {
        Load-InfrastructureConfig -Path $ofd.FileName
    }
})

$btnSaveLog.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = "Text files (*.txt)|*.txt"; $sfd.FileName = "Decom_Log_$(Get-Date -f 'yyyyMMdd').txt"
    if($sfd.ShowDialog() -eq "OK") { ($rtbLog.Text) | Out-File $sfd.FileName; Log-Msg "Log saved." "Cyan" }
})

$btnLogin.Add_Click({
    try { $Cred = Get-Credential; if($Cred){ $Script:RunCreds = $Cred; Log-Msg "Credentials loaded for session." "Cyan" } } catch { }
})

# --- DISCOVERY LOGIC (BATCH LOOP) ---
$btnCheck.Add_Click({
    $Targets = Get-TargetList
    if ($Targets.Count -eq 0) { return }

    $Script:IsPhysical = $chkPhys.Checked
    $Script:VmmTargetHost = @{} 
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # Reset State
    $btnStop.Enabled=$false; $btnDecom.Enabled=$false; $btnBackup.Enabled=$false; $btnDel.Enabled=$false
    $btnStop.BackColor=$Theme.BgControl; $btnDecom.BackColor=$Theme.BgControl; $btnDel.BackColor=$Theme.BgControl

    $Prefix = Get-TicketPrefix
    $FoundCount = 0

    foreach ($Target in $Targets) {
        Log-Msg "${Prefix}Processing Discovery for: $Target" "White"; [System.Windows.Forms.Application]::DoEvents()
        
        # 1. AD Discovery
        try {
            $Searcher = [ADSISearcher]""
            $Searcher.SearchRoot = [ADSI]"LDAP://$Script:TargetOU"
            $Searcher.Filter = "(&(objectClass=computer)(name=$Target))"
            $Result = $Searcher.FindOne()
            if ($Result) {
                $DN = $Result.Properties.distinguishedname[0]
                Log-Msg "${Prefix}  [AD] Verified: $DN" "Cyan"
                $FoundCount++
            } else { Log-Msg "${Prefix}  [AD] Not found in Target OU." "Red" }
        } catch { Log-Msg "${Prefix}  [AD] Error." "Red" }

        # 2. VMM Discovery (Skip if Physical)
        if (!$Script:IsPhysical) {
            $VmmFound = $false
            foreach ($Srv in $Script:VmmServers) {
                try {
                    $VM = Get-SCVirtualMachine -VMMServer $Srv -Name $Target -EA SilentlyContinue
                    if ($VM) {
                        Log-Msg "${Prefix}  [VMM] Found on host $($VM.VMHostName)." "Cyan"
                        $Script:VmmTargetHost[$Target] = $Srv
                        $VmmFound = $true; $FoundCount++
                        break
                    }
                } catch { }
            }
            if (!$VmmFound) { Log-Msg "${Prefix}  [VMM] Not found." "Red" }
        }
    }

    # ENABLE BUTTONS IF MACHINES ARE FOUND
    if ($FoundCount -gt 0) {
        $btnStop.Enabled = $true; $btnStop.BackColor = $Theme.AccentWarn
        $btnDecom.Enabled = $true; $btnDecom.BackColor = $Theme.AccentGreen
        $btnBackup.Enabled = $true; $btnBackup.BackColor = $Theme.AccentBlue
        
        # PHYSICAL LOGIC
        if ($Script:IsPhysical) {
            $btnDel.Enabled = $false
            $btnDel.BackColor = $Theme.BgControl
            Log-Msg "--- Discovery Complete (Physical Mode). 'Delete VM' is disabled. ---" "Yellow"
        } else {
            $btnDel.Enabled = $true
            $btnDel.BackColor = $Theme.AccentRed
            Log-Msg "--- Discovery Complete. Ready for Batch Actions. ---" "Lime"
        }

    } else {
        Log-Msg "--- No valid targets found. Check names or OU. ---" "Orange"
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# --- STOP ACTION ---
$btnStop.Add_Click({
    $Targets = Get-TargetList
    $Prefix = Get-TicketPrefix
    Log-Msg "${Prefix}--- STOPPING BATCH ($($Targets.Count) Servers) ---" "White"
    
    foreach ($Target in $Targets) {
        if ($Script:IsPhysical) {
            try { Stop-Computer -ComputerName $Target -Force -ErrorAction Stop; Log-Msg "${Prefix}[$Target] Shutdown signal sent." "Lime" } catch { Log-Msg "${Prefix}[$Target] WMI Error." "Red" }
        } else {
            $VMM = $Script:VmmTargetHost[$Target]
            if ($VMM) {
                try { Stop-SCVirtualMachine -VMMServer $VMM -VM $Target -Confirm:$false -EA Stop; Log-Msg "${Prefix}[$Target] VMM Shutdown complete." "Lime" } catch { Log-Msg "${Prefix}[$Target] VMM Error." "Red" }
            } else { Log-Msg "${Prefix}[$Target] Skip: Not mapped to VMM." "Yellow" }
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
})

# --- LOGICAL CLEANUP (CITRIX SAFE + BATCH) ---
$btnDecom.Add_Click({
    $Targets = Get-TargetList
    $Prefix = Get-TicketPrefix
    
    $Confirm = [System.Windows.Forms.MessageBox]::Show("Confirm LOGICAL CLEANUP on $($Targets.Count) servers?", "Confirm Batch", 4, 32)
    if ($Confirm -ne "Yes") { return }

    Log-Msg "${Prefix}--- STARTING BATCH CLEANUP ---" "White"
    
    foreach ($Target in $Targets) {
        # Ping Check
        if (Test-Connection -ComputerName $Target -Count 1 -Quiet) {
            Log-Msg "${Prefix}[$Target] SKIPPED: Server is still ONLINE." "Red"
            continue
        }

        # AD Cleanup
        try {
            $Searcher = [ADSISearcher]""
            $Searcher.SearchRoot = [ADSI]"LDAP://$Script:TargetOU"
            $Searcher.Filter = "(&(objectClass=computer)(name=$Target))"
            $Result = $Searcher.FindOne()
            if ($Result) {
                $DN = $Result.Properties.distinguishedname[0]
                $ADSIObj = [ADSI]"LDAP://$DN"
                $ADSIObj.DeleteTree()
                Log-Msg "${Prefix}[$Target] AD: Purged." "Lime"
            }
        } catch { Log-Msg "${Prefix}[$Target] AD: Error." "Red" }

        # SCCM Cleanup
        try {
            if (!(Test-Path "$($Script:SccmSiteCode):")) { New-PSDrive -Name $Script:SccmSiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root $Script:SccmProvider -EA 0 | Out-Null }
            $Current = (Get-Location).Path; Set-Location "$($Script:SccmSiteCode):"
            Remove-CMDevice -DeviceName $Target -Force -Confirm:$false -EA 0
            Set-Location $Current
            Log-Msg "${Prefix}[$Target] SCCM: Removed." "Lime"
        } catch { Log-Msg "${Prefix}[$Target] SCCM: Error/Not Found." "Yellow"; Set-Location "C:" }

        # CITRIX CLEANUP (SAFE MODE)
        if ($chkCitrix.Checked) {
            if ($Script:IsPhysical) {
                Log-Msg "${Prefix}[$Target] Citrix: Skipped (Physical Mode)." "Yellow"
            } else {
                try {
                    Log-Msg "${Prefix}[$Target] Citrix: Attempting removal..." "White"
                    # Add-PSSnapin Citrix* -EA SilentlyContinue
                    # Remove-BrokerMachine -MachineName $Target -Force -ErrorAction Stop
                    Log-Msg "${Prefix}[$Target] Citrix: Removed (Simulated/Placeholder)." "Lime"
                } catch {
                    Log-Msg "${Prefix}[$Target] Citrix: Machine not found or Error (Non-fatal)." "Yellow"
                }
            }
        }

        # DNS Cleanup
        try {
            Remove-DnsServerResourceRecord -ComputerName $Script:DC -ZoneName $Script:ForwardZone -Name $Target -RRType "A" -Force -EA 0
            Log-Msg "${Prefix}[$Target] DNS: Removed." "Lime"
        } catch { }
        
        [System.Windows.Forms.Application]::DoEvents()
    }
})

# --- DELETE FROM HYPERVISOR (BLOCKED IF PHYSICAL) ---
$btnDel.Add_Click({
    $Targets = Get-TargetList
    $Prefix = Get-TicketPrefix
    
    # DOUBLE SAFETY CHECK
    if ($Script:IsPhysical) { 
        [System.Windows.Forms.MessageBox]::Show("Operation not allowed in Physical Mode.", "Blocked", 0, 16)
        Log-Msg "${Prefix}ABORTED: Physical hardware protection." "Red"
        return 
    }
    
    $InputUser = [Microsoft.VisualBasic.Interaction]::InputBox("DANGER: PERMANENTLY DELETING $($Targets.Count) VMs.`nType 'DELETE' to confirm:", "Confirm Destructive Action", "")
    
    if ($InputUser -ne "DELETE") { return }

    Log-Msg "${Prefix}--- STARTING BATCH DISK DELETION ---" "Red"
    
    foreach ($Target in $Targets) {
        $VMM = $Script:VmmTargetHost[$Target]
        if ($VMM) {
            Log-Msg "${Prefix}[$Target] Deleting from Storage..." "Orange"
            try { 
                Remove-SCVirtualMachine -VMMServer $VMM -VM $Target -Force -Confirm:$false
                Log-Msg "${Prefix}[$Target] DELETED PERMANENTLY." "Red" 
            } catch { 
                Log-Msg "${Prefix}[$Target] VMM Error: $($_.Exception.Message)" "Red" 
            }
        } else {
            Log-Msg "${Prefix}[$Target] Skip: VMM Host unknown." "Yellow"
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
})

# --- AUTO LOAD CONFIG AT STARTUP ---
# Try to load local config automatically on launch
if ([string]::IsNullOrEmpty($PSScriptRoot)) { $BaseDir = (Get-Location).Path } else { $BaseDir = $PSScriptRoot }
$AutoXml = Join-Path $BaseDir "DecomConfig.xml"
if (Test-Path $AutoXml) { Load-InfrastructureConfig $AutoXml } else { Log-Msg "Info: DecomConfig.xml not found. Using defaults. Load manually if needed." "Yellow" }

$form.ShowDialog() | Out-Null; $form.Dispose()
