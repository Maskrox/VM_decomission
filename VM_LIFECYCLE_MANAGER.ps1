# =============================================================================
# VM LIFECYCLE MANAGER - ENTERPRISE EDITION (V6.1)
#
# This script provides a graphical user interface (GUI) to automate the
# decommissioning process for virtual machines (VMs) and physical servers
# across various infrastructure components such as Active Directory, DNS,
# System Center Virtual Machine Manager (SCVMM), System Center Configuration
# Manager (SCCM), and Citrix Virtual Apps and Desktops.
#
# This version supports parallel processing of decommissioning tasks using
# PowerShell Jobs (DecomWorker.ps1) for improved performance.
#
# Date: February 2026
# Version: 6.2 
# =============================================================================

# --- FIXED LOGGING FUNCTION ---
# This function is used for logging messages to the RichTextBox on the GUI.
# It ensures thread-safe updates to the GUI by using $rtbLog.Invoke if necessary.
function Log-Msg {
    param (
        [string]$Msg,         # The message string to log.
        [string]$ColorName="Lime" # The color of the message (e.g., "Lime", "Red", "Yellow", "Cyan").
    )

    # Check if $rtbLog (RichTextBox) is initialized and available
    if ($rtbLog -and $rtbLog.GetType().Name -eq "RichTextBox" -and $rtbLog.IsHandleCreated) {
        # Original GUI logging logic
        $Action = {
            $rtbLog.SelectionStart = $rtbLog.TextLength; # Set cursor to the end of text.
            $rtbLog.SelectionColor = [System.Drawing.Color]::FromName($ColorName) # Set text color.
            $rtbLog.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $Msg`r`n"); # Append message with timestamp.
            $rtbLog.ScrollToCaret() # Scroll to the newly added text.
        }

        if ($rtbLog.InvokeRequired) {
            $rtbLog.Invoke($Action)
        } else {
            & $Action
        }
    } else {
        # Fallback to console logging if RichTextBox is not ready
        $timestamp = Get-Date -Format 'HH:mm:ss'
        $formattedMsg = "[$timestamp] $Msg"
        
        # Determine console color based on $ColorName (basic mapping)
        switch ($ColorName.ToLower()) {
            "red"    { Write-Host -Object $formattedMsg -ForegroundColor Red }
            "green"  { Write-Host -Object $formattedMsg -ForegroundColor Green }
            "yellow" { Write-Host -Object $formattedMsg -ForegroundColor Yellow }
            "cyan"   { Write-Host -Object $formattedMsg -ForegroundColor Cyan }
            "orange" { Write-Host -Object $formattedMsg -ForegroundColor DarkYellow } # PowerShell doesn't have "Orange"
            "lime"   { Write-Host -Object $formattedMsg -ForegroundColor Green } # Lime is often represented as Green
            "white"  { Write-Host -Object $formattedMsg -ForegroundColor White }
            "gray"   { Write-Host -Object $formattedMsg -ForegroundColor Gray }
            default  { Write-Host -Object $formattedMsg -ForegroundColor White } # Default to white
        }
    }
}
# --- 1. HIGH DPI AWARENESS ---
# Attempts to enable High DPI awareness for the PowerShell process.
# This helps ensure the GUI scales correctly on high-resolution displays.
try {
    $code = '[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();'
    $Win32 = Add-Type -MemberDefinition $code -Name "Win32" -Namespace "Win32" -PassThru
    $Win32::SetProcessDPIAware() | Out-Null
} catch {
    # Suppress errors if setting DPI awareness fails (e.g., on older OS versions)
}

# --- LOAD ASSEMBLIES & MODULES ---
# Load essential .NET assemblies for GUI components.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic # Used for InputBox

# Import core PowerShell modules required for infrastructure management.
# These modules are needed in the main GUI script for discovery phases.
Import-Module ActiveDirectory, VirtualMachineManager -ErrorAction SilentlyContinue

# Citrix Modules/Snapins - Depending on XenApp/XenDesktop version, one or more may be needed.
# This section attempts to load Citrix PowerShell components, primarily for discovery.
# The worker script (DecomWorker.ps1) will also load these modules as needed.
$citrixSnapinsLoaded = $false
try {
    $loadedSnapins = Add-PSSnapin Citrix* -PassThru -ErrorAction SilentlyContinue
    if ($loadedSnapins) {
        $citrixSnapinsLoaded = $true
        Log-Msg "Citrix PowerShell Snapins loaded successfully (Main Script)." "Green"
    }
} catch {
    # This catch block would only be hit if Add-PSSnapin throws a terminating error.
}

if (-not $citrixSnapinsLoaded) {
    Log-Msg "Citrix PowerShell Snapins not found/loaded (Main Script), attempting Import-Module for newer SDKs." "Yellow"
    $citrixModulesLoaded = $false
    try {
        # Attempt to import common Citrix modules for modern SDKs
        $loadedModules = Import-Module Broker, Configuration, Host, DelegatedAdmin -PassThru -ErrorAction SilentlyContinue
        if ($loadedModules) {
            $citrixModulesLoaded = $true
            Log-Msg "Citrix PowerShell Modules loaded successfully (Main Script)." "Green"
        }
    } catch {
        # This catch block would only be hit if Import-Module throws a terminating error.
    }
    if (-not $citrixModulesLoaded) {
        Log-Msg "Citrix PowerShell Modules also failed to load (Main Script). Citrix discovery functionality may be limited." "Red"
    }
}

# --- GLOBAL VARIABLES (INITIALIZE EMPTY) ---
# These variables are script-scoped and hold state information throughout the application's lifecycle.
$Script:RunCreds = $null               # Stores user credentials for authenticated operations
$Script:VmmTargetHost = @{}             # Maps target VM names to their VMM servers during discovery
$Script:IsPhysical = $false             # Flag to indicate if the target is physical hardware

# Infrastructure Variables (Script Scope for dynamic updates from DecomConfig.xml)
# These are populated from the configuration file and used by various cleanup functions.
$Script:VmmServers    = @()             # List of VMM server names
$Script:DC            = ""              # Domain Controller for DNS operations
$Script:SccmSiteCode  = ""              # SCCM Site Code
$Script:SccmProvider  = ""              # SCCM SMS Provider server
$Script:ForwardZone   = ""              # DNS Forward Lookup Zone for cleanup
$Script:TargetOU      = ""              # Target Active Directory OU for computer object lookup/deletion
$Script:CitrixDDC     = ""              # Citrix Delivery Controller (currently unused, but reserved)
$Script:AllowedADGroup = ""              # Active Directory group for authentication
$Script:AuditLogPath = ""              # Path to the audit log file
$Script:SmtpServer = ""                # SMTP Server for email notifications
$Script:SmtpPort = 0                   # SMTP Port for email notifications
$Script:SenderEmail = ""               # Sender email address for notifications
$Script:RecipientEmails = @()          # Recipient email addresses (array)
$Script:EnableSsl = $false             # Enable SSL for SMTP
$Script:SmtpUsername = ""              # Username for SMTP authentication (if required)
$Script:SmtpPassword = $null           # Password for SMTP authentication (if required, stored securely)

# --- THEME DEFINITION (HARDCODED DEFAULT) ---
# Defines the color scheme for the GUI elements, enhancing readability and visual appeal.
$Theme = @{
    BgBase      = [System.Drawing.Color]::FromArgb(32, 32, 32)    # Main background color
    BgControl   = [System.Drawing.Color]::FromArgb(45, 45, 48)    # Background for controls (buttons, textboxes)
    TextMain    = [System.Drawing.Color]::FromArgb(240, 240, 240) # Primary text color
    TextDim     = [System.Drawing.Color]::FromArgb(160, 160, 160) # Dimmed text color for descriptions
    AccentBlue  = [System.Drawing.Color]::FromArgb(0, 120, 215)   # Accent color for titles, highlights
    AccentGreen = [System.Drawing.Color]::FromArgb(16, 124, 16)   # Color for success messages
    AccentRed   = [System.Drawing.Color]::FromArgb(232, 17, 35)   # Color for error/destructive actions
    AccentWarn  = [System.Drawing.Color]::FromArgb(255, 140, 0)   # Color for warnings
    Border      = [System.Drawing.Color]::FromArgb(80, 80, 80)    # Border color
}

# Font definitions for various text elements within the GUI.
$FontTitle = New-Object System.Drawing.Font("Segoe UI Semibold", 18)
$FontHead  = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$FontText  = New-Object System.Drawing.Font("Segoe UI", 10)
$FontSmall = New-Object System.Drawing.Font("Segoe UI", 8)
$FontMono  = New-Object System.Drawing.Font("Consolas", 9)

# --- UI SETUP ---
# Initializes the main Windows Form for the application.
$form = New-Object System.Windows.Forms.Form
$form.Text = "Server Decommissioning Manager - Enterprise Master"; # Window title
$form.Size = New-Object System.Drawing.Size(950, 950)              # Initial window size
$form.MinimumSize = New-Object System.Drawing.Size(950, 800)       # Minimum allowable window size
$form.BackColor = $Theme.BgBase                                    # Background color from theme
$form.StartPosition = "CenterScreen"                               # Center window on screen
$form.FormBorderStyle = "Sizable"                                  # Allow resizing
$form.MaximizeBox = $true                                          # Show maximize button

# Main Panel with Scroll Support: Hosts all other controls and provides scrolling if content overflows.
$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Dock = "Fill"                                           # Fills the entire form
$mainPanel.AutoScroll = $true                                      # Enables auto-scrolling
$mainPanel.AutoScrollMinSize = New-Object System.Drawing.Size(900, 1015) # Minimum scrollable area
$form.Controls.Add($mainPanel)                                     # Add panel to the form

# --- UI HELPER FUNCTIONS ---
# Functions to simplify the creation of styled GUI elements, promoting consistency.

function New-StyledButton {
    param (
        [string]$Text,      # Text displayed on the button
        [System.Drawing.Color]$Color, # Background color of the button
        [int]$X,            # X-coordinate of the button's top-left corner
        [int]$Y,            # Y-coordinate of the button's top-left corner
        [int]$W,            # Width of the button
        [int]$H,            # Height of the button
        [System.Windows.Forms.Control]$Parent, # Parent control to add the button to
        [System.Windows.Forms.AnchorStyles]$Anchor # Anchor style for resizing behavior
    )
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, $Y)
    $btn.Size = New-Object System.Drawing.Size($W, $H)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    if ($Color) { $btn.BackColor = $Color } else { $btn.BackColor = $Theme.BgControl }
    $btn.ForeColor = $Theme.TextMain
    $btn.Font = $FontHead
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand # Changes cursor on hover
    if ($Anchor) { $btn.Anchor = $Anchor }
    $Parent.Controls.Add($btn)
    return $btn
}

function New-StyledGroup {
    param (
        [string]$Text,      # Title text for the group box
        [int]$X,            # X-coordinate of the group box's top-left corner
        [int]$Y,            # Y-coordinate of the group box's top-left corner
        [int]$W,            # Width of the group box
        [int]$H,            # Height of the group box
        [System.Windows.Forms.Control]$Parent # Parent control to add the group box to
    )
    $grp = New-Object System.Windows.Forms.GroupBox
    $grp.Text = $Text
    $grp.Location = New-Object System.Drawing.Point($X, $Y)
    $grp.Size = New-Object System.Drawing.Size($W, $H)
    $grp.ForeColor = $Theme.AccentBlue # Title color from theme
    $grp.Font = $FontHead
    # Anchors to top, left, and right to expand with the parent panel
    $grp.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $Parent.Controls.Add($grp)
    return $grp
}

function New-DescLabel {
    param (
        [string]$Text,      # Text for the description label
        [int]$X,            # X-coordinate
        [int]$Y,            # Y-coordinate
        [int]$W,            # Width
        [System.Windows.Forms.Control]$Parent # Parent control
    )
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.Location = New-Object System.Drawing.Point($X, $Y)
    $lbl.Size = New-Object System.Drawing.Size($W, 20) # Fixed height
    $lbl.ForeColor = $Theme.TextDim                    # Dimmed text color from theme
    $lbl.Font = $FontSmall
    $lbl.TextAlign = "TopCenter"                       # Text alignment
    $Parent.Controls.Add($lbl)
}

# --- GUI CONSTRUCTION ---
# This section defines and places all graphical elements on the form.

# Header Title: Displays the main title of the application.
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "VM LIFECYCLE MANAGER"
$lblTitle.Font = $FontTitle
$lblTitle.ForeColor = $Theme.AccentBlue
$lblTitle.Location = New-Object System.Drawing.Point(20, 20)
$lblTitle.Size = New-Object System.Drawing.Size(400, 40)
$mainPanel.Controls.Add($lblTitle)

# --- BUTTON: LOAD CONFIG ---
# Button to manually load the infrastructure configuration XML file.
$btnLoadConfig = New-StyledButton -Text "LOAD CONFIG" -Color $Theme.BgControl -X 440 -Y 20 -W 180 -H 35 -Parent $mainPanel -Anchor ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$btnLoadConfig.FlatAppearance.BorderColor = $Theme.TextDim
$btnLoadConfig.FlatAppearance.BorderSize = 1
$btnLoadConfig.Font = $FontText

# Login Button: Prompts the user for credentials to be used for various operations.
$btnLogin = New-StyledButton -Text "LOGIN CREDENTIALS" -Color $Theme.AccentBlue -X 640 -Y 20 -W 240 -H 35 -Parent $mainPanel -Anchor ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
$btnLogin.Font = $FontText

# Label to display the currently logged-in user
$lblLoggedInUser = New-Object System.Windows.Forms.Label
$lblLoggedInUser.Text = "Not Logged In" # Initial placeholder
$lblLoggedInUser.Location = New-Object System.Drawing.Point(640, 60) # Position below login button
$lblLoggedInUser.Size = New-Object System.Drawing.Size(240, 20)
$lblLoggedInUser.ForeColor = $Theme.TextDim
$lblLoggedInUser.Font = $FontSmall
$lblLoggedInUser.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$mainPanel.Controls.Add($lblLoggedInUser)

# GROUP 1: DISCOVERY ZONE - Allows identification of target machines and initial settings.
$grpDisc = New-StyledGroup -Text " Phase 1: Identification (Batch Support) " -X 20 -Y 70 -W 880 -H 140 -Parent $mainPanel

# Label for target hostnames input.
$lblVM = New-Object System.Windows.Forms.Label
$lblVM.Text = "Target Hostnames (comma separated):" 
$lblVM.Location = New-Object System.Drawing.Point(20, 40); $lblVM.Size = New-Object System.Drawing.Size(300, 25)
$lblVM.ForeColor = $Theme.TextDim; $lblVM.Font = $FontText
$grpDisc.Controls.Add($lblVM)

# Textbox for entering one or more target VM hostnames.
$txtVM = New-Object System.Windows.Forms.TextBox
$txtVM.Location = New-Object System.Drawing.Point(20, 65); $txtVM.Size = New-Object System.Drawing.Size(350, 28)
$txtVM.Font = New-Object System.Drawing.Font("Segoe UI", 11); $txtVM.BackColor = $Theme.BgControl; $txtVM.ForeColor = $Theme.TextMain; $txtVM.BorderStyle = "FixedSingle"
$txtVM.Enabled = $false # Disable until login
$grpDisc.Controls.Add($txtVM)

# Label for ticket number input.
$lblTick = New-Object System.Windows.Forms.Label
$lblTick.Text = "Ticket / Task #:" 
$lblTick.Location = New-Object System.Drawing.Point(380, 40); $lblTick.Size = New-Object System.Drawing.Size(150, 25)
$lblTick.ForeColor = $Theme.TextDim; $lblTick.Font = $FontText
$grpDisc.Controls.Add($lblTick)

# Textbox for entering a ticket or task number (optional, for logging).
$txtTicket = New-Object System.Windows.Forms.TextBox
$txtTicket.Location = New-Object System.Drawing.Point(380, 65); $txtTicket.Size = New-Object System.Drawing.Size(140, 28)
$txtTicket.Font = New-Object System.Drawing.Font("Segoe UI", 11); $txtTicket.BackColor = $Theme.BgControl; $txtTicket.ForeColor = $Theme.AccentWarn; $txtTicket.BorderStyle = "FixedSingle"
$txtTicket.Enabled = $false # Disable until login
$grpDisc.Controls.Add($txtTicket)

# Button to initiate the auto-discovery process for the entered hostnames.
$btnCheck = New-StyledButton -Text "1. AUTO-DISCOVER" -Color $Theme.BgControl -X 540 -Y 63 -W 180 -H 32 -Parent $grpDisc
$btnCheck.FlatAppearance.BorderColor = $Theme.AccentBlue; $btnCheck.FlatAppearance.BorderSize = 1
$btnCheck.Enabled = $false # Disable until login

# Checkbox to specify if the target is physical hardware, which affects available actions.
$chkPhys = New-Object System.Windows.Forms.CheckBox
$chkPhys.Text = "Physical Hardware"
$chkPhys.Location = New-Object System.Drawing.Point(20, 100); $chkPhys.Size = New-Object System.Drawing.Size(240, 25)
$chkPhys.ForeColor = $Theme.TextMain; $chkPhys.Font = $FontText
$chkPhys.Enabled = $false # Disable until login
$grpDisc.Controls.Add($chkPhys)

# Checkbox to enable/disable Citrix cleanup operations.
$chkCitrix = New-Object System.Windows.Forms.CheckBox
$chkCitrix.Text = "Include Citrix Cleanup"
$chkCitrix.Location = New-Object System.Drawing.Point(280, 100); $chkCitrix.Size = New-Object System.Drawing.Size(240, 25)
$chkCitrix.ForeColor = $Theme.AccentWarn; $chkCitrix.Font = $FontText
$chkCitrix.Enabled = $false # Disable until login
$grpDisc.Controls.Add($chkCitrix)

# Display Area for Discovered Targets (ListView)
$lvTargets = New-Object System.Windows.Forms.ListView
$lvTargets.Location = New-Object System.Drawing.Point(20, 220) # Position below $grpDisc
$lvTargets.Size = New-Object System.Drawing.Size(880, 150)    # Adjust size as needed
$lvTargets.View = [System.Windows.Forms.View]::Details        # Show column headers
$lvTargets.FullRowSelect = $true                               # Select entire row
$lvTargets.GridLines = $true                                   # Show grid lines
$lvTargets.MultiSelect = $false                                # Allow single selection
$lvTargets.BackColor = $Theme.BgControl
$lvTargets.ForeColor = $Theme.TextMain
$lvTargets.Font = $FontText
$lvTargets.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$mainPanel.Controls.Add($lvTargets)

# Define columns for the ListView
$colHost = New-Object System.Windows.Forms.ColumnHeader
$colHost.Text = "Hostname"
$colHost.Width = 150
$lvTargets.Columns.Add($colHost)

$colAD = New-Object System.Windows.Forms.ColumnHeader
$colAD.Text = "AD Status"
$colAD.Width = 100
$lvTargets.Columns.Add($colAD)

$colVMM = New-Object System.Windows.Forms.ColumnHeader
$colVMM.Text = "VMM Status"
$colVMM.Width = 100
$lvTargets.Columns.Add($colVMM)

$colOverall = New-Object System.Windows.Forms.ColumnHeader
$colOverall.Text = "Overall Status"
$colOverall.Width = 120
$lvTargets.Columns.Add($colOverall)

$colLastAction = New-Object System.Windows.Forms.ColumnHeader
$colLastAction.Text = "Last Action"
$colLastAction.Width = 280 # Wider for action details
$lvTargets.Columns.Add($colLastAction)

# GROUP 2: EXECUTION ZONE - Contains buttons to initiate various decommissioning phases.
$grpAct = New-StyledGroup -Text " Phase 2: Batch Execution Sequence " -X 20 -Y 390 -W 880 -H 150 -Parent $mainPanel

# Buttons for decommissioning actions. These are enabled after successful discovery.
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

# LOGS AREA - Displays real-time activity and allows saving the log.
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Process Activity Log:"
$lblLog.Location = New-Object System.Drawing.Point(25, 400); $lblLog.Size = New-Object System.Drawing.Size(200, 25)
$lblLog.ForeColor = $Theme.TextDim; $lblLog.Font = $FontText
$mainPanel.Controls.Add($lblLog)

# Search controls for activity log
$txtLogSearch = New-Object System.Windows.Forms.TextBox
$txtLogSearch.Location = New-Object System.Drawing.Point(25, 430)
$txtLogSearch.Size = New-Object System.Drawing.Size(750, 28)
$txtLogSearch.Font = $FontText
$txtLogSearch.BackColor = $Theme.BgControl
$txtLogSearch.ForeColor = $Theme.TextMain
$txtLogSearch.BorderStyle = "FixedSingle"

$txtLogSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$mainPanel.Controls.Add($txtLogSearch)

$btnLogSearch = New-StyledButton -Text "SEARCH" -Color $Theme.BgControl -X 785 -Y 428 -W 115 -H 32 -Parent $mainPanel
$btnLogSearch.FlatAppearance.BorderColor = $Theme.AccentBlue
$btnLogSearch.FlatAppearance.BorderSize = 1
$btnLogSearch.Font = $FontText
$btnLogSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$mainPanel.Controls.Add($btnLogSearch)


$rtbLog = New-Object System.Windows.Forms.RichTextBox
$rtbLog.Location = New-Object System.Drawing.Point(25, 465); $rtbLog.Size = New-Object System.Drawing.Size(875, 380) # Adjusted Y and Height
$rtbLog.BackColor = "Black"; $rtbLog.ForeColor = $Theme.AccentGreen; $rtbLog.Font = $FontMono; $rtbLog.ReadOnly = $true; $rtbLog.BorderStyle = "None"
$rtbLog.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$mainPanel.Controls.Add($rtbLog)

# Button to save the contents of the activity log to a text file.
$btnSaveLog = New-StyledButton -Text "SAVE LOG TO FILE" -Color $Theme.BgControl -X 25 -Y 900 -W 200 -H 35 -Parent $mainPanel -Anchor ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
$btnSaveLog.Font = $FontText; $btnSaveLog.FlatAppearance.BorderColor = $Theme.AccentBlue; $btnSaveLog.FlatAppearance.BorderSize = 1

# Label for detailed progress text
$lblProgressText = New-Object System.Windows.Forms.Label
$lblProgressText.Text = "Idle" # Initial text
$lblProgressText.Location = New-Object System.Drawing.Point(25, 840)
$lblProgressText.Size = New-Object System.Drawing.Size(875, 20)
$lblProgressText.ForeColor = $Theme.TextMain
$lblProgressText.Font = $FontText
$lblProgressText.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblProgressText.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$mainPanel.Controls.Add($lblProgressText)

# Progress Bar for Batch Operations
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(25, 865) # Adjusted Y
$progressBar.Size = New-Object System.Drawing.Size(875, 15)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$mainPanel.Controls.Add($progressBar)

# =============================================================================

# BUSINESS LOGIC & FUNCTIONS

# This section contains the core functionality and helper functions of the script.

# =============================================================================



# --- HELPER FUNCTION: Test-ADGroupMembership ---

# Checks if a given user is a member of a specified Active Directory group.

function Test-ADGroupMembership {

    param (

        [Parameter(Mandatory=$true)]

        [string]$UserName,

        [Parameter(Mandatory=$true)]

        [string]$GroupName

    )



    # Note: Assumes ActiveDirectory module is loaded.

    # We do NOT use the $Credential parameter here for Get-ADUser/Group

    # because these cmdlets typically run in the context of the calling user

    # or require explicit authentication which is handled by the initial Get-Credential prompt.

    # Trying to pass PSCredential directly to Get-ADUser can be complex for validation

    # without a specific domain context or trust.

    # The primary goal is to validate if the *currently authenticated user* (via Get-Credential)

    # is a member of the AD group.



    try {

        # Get the user object using the sAMAccountName (assuming the username from Get-Credential is sAMAccountName)

        # Or could be UserPrincipalName, Get-ADUser is flexible.

        $user = Get-ADUser -Identity $UserName -Properties MemberOf -ErrorAction Stop

        

        # Get the group object

        $group = Get-ADGroup -Identity $GroupName -ErrorAction Stop



        # Check if the user is a member of the group

        # The MemberOf property contains an array of Distinguished Names of groups the user is directly a member of.

        $isMember = $user.MemberOf -contains $group.DistinguishedName



        return $isMember



    } catch {

                $errorMessage = ($_.Exception).Message

                Log-Msg ("AD Group Membership check failed for " + $UserName + " in " + $GroupName + ": " + $errorMessage) "Red"

                Audit-Log -Action "AD Group Membership Check" -Target $GroupName -Outcome "Error" -Details ("Failed for user " + $UserName + ": " + $errorMessage) -User $UserName

                return $false

    }

}


# --- HELPER FUNCTION: Audit-Log ---
# Records an audit entry to the configured audit log file.
function Audit-Log {
    param (
        [string]$User = "(System)", # User performing the action (defaults to System if not specified)
        [Parameter(Mandatory=$true)]
        [string]$Action,      # Description of the action performed
        [string]$Target = "N/A",      # Target of the action (e.g., VM name, config file)
        [string]$Outcome = "Info",    # Result of the action (e.g., Success, Failure, Info)
        [string]$Details = ""         # Additional details or error messages
    )

    if (-not $Script:AuditLogPath) {
        Log-Msg "Audit-Log: AuditLogPath is not configured. Skipping audit log entry." "Yellow"
        return
    }

    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timestamp | User: $User | Action: $Action | Target: $Target | Outcome: $Outcome | Details: $Details"

        # Ensure the directory exists
        $logDirectory = Split-Path -Path $Script:AuditLogPath -Parent
        if (-not (Test-Path $logDirectory)) {
            New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        }

        Add-Content -Path $Script:AuditLogPath -Value $logEntry
        # Log to GUI as well, but with a less prominent color
        Log-Msg "AUDIT: $logEntry" "Gray" 

    } catch {
        Log-Msg "Audit-Log: Failed to write to audit log file '$Script:AuditLogPath': $($_.Exception.Message)" "Red"
    }
}


# --- HELPER FUNCTION: Send-DecomEmail ---
# Sends an email notification using configured settings.
function Send-DecomEmail {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        [Parameter(Mandatory=$true)]
        [string]$Body,
        [bool]$IsHtml = $false
    )

    if (-not $Script:SmtpServer -or -not $Script:SenderEmail -or -not $Script:RecipientEmails) {
        Log-Msg "Send-DecomEmail: Email settings are incomplete (SmtpServer, SenderEmail, or RecipientEmails missing). Skipping email notification." "Yellow"
        Audit-Log -Action "Send Email" -Target "N/A" -Outcome "Warning" -Details "Email settings incomplete. Subject: $Subject" -User "(System)"
        return
    }

    try {
        $sendMailParams = @{
            SmtpServer = $Script:SmtpServer
            From       = $Script:SenderEmail
            To         = ($Script:RecipientEmails -split ',' | ForEach-Object { $_.Trim() }) # Split and trim recipients
            Subject    = $Subject
            Body       = $Body
            UseSSL     = $Script:EnableSsl
            Port       = $Script:SmtpPort
            BodyAsHtml = $IsHtml
        }

        # Handle SMTP authentication if username is provided
        if (-not [string]::IsNullOrEmpty($Script:SmtpUsername)) {
            # Ideally, password would be securely fetched here (e.g., from Credential Manager)
            # For now, we'll assume integrated auth or password already loaded if username is present.
            # If $Script:SmtpPassword is a PSCredential object, use it. Otherwise, prompt or use system creds.
            # For this context, we will not automatically prompt for password here, expecting it to be available if needed.
            # If $Script:SmtpPassword isn't set, then Send-MailMessage will likely fail without user interaction.
            if ($Script:SmtpPassword) {
                $sendMailParams.Add("Credential", $Script:SmtpPassword)
            } else {
                Log-Msg "Send-DecomEmail: SMTP Username is configured but no password provided. Attempting without explicit password." "Yellow"
                Audit-Log -Action "Send Email" -Target "N/A" -Outcome "Warning" -Details "SMTP Username provided but no password for $Subject." -User "(System)"
            }
        }

        Send-MailMessage @sendMailParams -ErrorAction Stop

        Log-Msg "Email sent successfully with subject: '$Subject'." "Lime"
        Audit-Log -Action "Send Email" -Target "Recipients: $($Script:RecipientEmails)" -Outcome "Success" -Details "Email sent: $Subject" -User "(System)"

    } catch {
        $errorMessage = ($_.Exception).Message
        Log-Msg "Send-DecomEmail: Failed to send email with subject '$Subject': $errorMessage" "Red"
        Audit-Log -Action "Send Email" -Target "Recipients: $($Script:RecipientEmails)" -Outcome "Failure" -Details "Failed to send email: $errorMessage. Subject: $Subject" -User "(System)"
    }
}


# --- Get Ticket Prefix Function ---
# Generates a log prefix based on the user-provided ticket number.
function Get-TicketPrefix {
    if ($txtTicket.Text.Trim()) { 
        return "[$($txtTicket.Text.Trim())] " 
    }
    return "" # Return empty string if no ticket is provided.
}

# --- Get Target List Function ---
# Parses the comma-separated string of hostnames from the input textbox.
function Get-TargetList {
    return $txtVM.Text -split "," | # Split the input string by commas.
           ForEach-Object { $_.Trim() } | # Trim whitespace from each hostname.
           Where-Object { $_ -ne "" }    # Filter out any empty entries.
}

# --- HELPER FUNCTION: Update-ListViewTargetStatus ---
# Updates the status of a target in the ListView.
function Update-ListViewTargetStatus {
    param (
        [Parameter(Mandatory=$true)]
        [string]$TargetName,
        [Parameter(Mandatory=$true)]
        [string]$OverallStatus,
        [Parameter(Mandatory=$true)]
        [string]$LastAction
    )
    
    # Find the ListViewItem for the given target
    $lvi = $lvTargets.FindItemWithText($TargetName, $false, 0, $true) # Find by hostname in first column

    if ($lvi) {
        $lvi.SubItems[3].Text = $OverallStatus # Overall Status (index 3)
        $lvi.SubItems[4].Text = $LastAction    # Last Action (index 4)
        
        # NEW: Set ForeColor based on OverallStatus
        switch ($OverallStatus) {
            "Ready for Action"   { $lvi.ForeColor = $Theme.TextMain } # Default color
            "Stopped"            { $lvi.ForeColor = $Theme.AccentWarn } # Warning/Action taken
            "Cleaned"            { $lvi.ForeColor = $Theme.AccentGreen } # Success
            "DELETED"            { $lvi.ForeColor = $Theme.AccentRed }   # Critical action
            "Not Found (Cannot Decom)" { $lvi.ForeColor = $Theme.TextDim } # Dimmed/Excluded
            "Stop Failed"        { $lvi.ForeColor = $Theme.AccentRed }
            "Clean Failed"       { $lvi.ForeColor = $Theme.AccentRed }
            "Delete Failed"      { $lvi.ForeColor = $Theme.AccentRed }
            default              { $lvi.ForeColor = $Theme.TextMain }   # Default to main text color
        }
        
        [System.Windows.Forms.Application]::DoEvents() # Update GUI immediately
    }
}

# --- HELPER FUNCTION: Update-ProgressLabel ---
# Updates the text of the progress label and ensures GUI responsiveness.
function Update-ProgressLabel {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message
    )
    $lblProgressText.Text = $Message
    [System.Windows.Forms.Application]::DoEvents()
}


# --- CONFIG LOADING LOGIC (FUNCTION WRAPPED) ---
# Loads infrastructure configuration from an XML file and populates global script variables.
function Load-InfrastructureConfig ($Path) {
    if (Test-Path $Path) {
        try {
            [xml]$XmlConfig = Get-Content $Path # Read and parse the XML file.
            $Infra = $XmlConfig.Configuration.Infrastructure # Navigate to the Infrastructure section.
            
            # Update Script Global Variables with values from the XML configuration.
            $Script:VmmServers    = @($Infra.VmmServers.Server)     # Array of VMM server names.
            $Script:DC            = $Infra.DomainController         # Primary Domain Controller for DNS.
            $Script:SccmSiteCode  = $Infra.SccmSiteCode             # SCCM Site Code.
            $Script:SccmProvider  = $Infra.SccmProvider             # SCCM SMS Provider server.
            $Script:ForwardZone   = $Infra.DnsZone                  # DNS Forward Lookup Zone.
            $Script:TargetOU      = $Infra.TargetOU                 # Active Directory OU for computer objects.
            $Script:CitrixDDC     = $Infra.CitrixController         # Citrix Delivery Controller.
            $Script:AllowedADGroup = $XmlConfig.Configuration.Security.AllowedADGroup # Load Allowed AD Group.
            $Script:AuditLogPath = $XmlConfig.Configuration.Security.AuditLogPath # Load Audit Log Path.
            
            Log-Msg "--- Configuration Loaded Successfully from: $Path ---" "Cyan"
            Audit-Log -Action "Configuration Load" -Target $Path -Outcome "Success" -Details "Loaded config from $Path" -User "(System)"
            Log-Msg "   > Target OU: $Script:TargetOU" "Cyan"
            Log-Msg "   > VMM Servers: $($Script:VmmServers -join ', ')" "Cyan"
            Log-Msg "   > Allowed AD Group: $Script:AllowedADGroup" "Cyan"
            Log-Msg "   > Audit Log Path: $Script:AuditLogPath" "Cyan"
            return $true # Indicate successful loading.
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error parsing XML. Check format.", "Config Error", 0, 16)
            $errorMessage = ($_.Exception).Message # Capture error message once
            Log-Msg "Failed to load config: $errorMessage" "Red" # Log the parsing error.
            Audit-Log -Action "Configuration Load" -Target $Path -Outcome "Failure" -Details "Error: $errorMessage" -User "(System)"
            return $false # Indicate failure.
        }
    } else {
        Log-Msg "Info: DecomConfig.xml not found. Using defaults. Load manually if needed." "Yellow"
        Audit-Log -Action "Configuration Load" -Target $AutoXml -Outcome "Info" -Details "DecomConfig.xml not found at startup." -User "(System)"
        return $false # File not found.
    }
}

# --- BUTTON EVENT HANDLERS ---

# This section defines the actions taken when various buttons on the GUI are clicked.



# 1. Load Config Button Action: Handles loading an external XML configuration file.

$btnLoadConfig.Add_Click({

    $ofd = New-Object System.Windows.Forms.OpenFileDialog # Create a new Open File Dialog.

    $ofd.Filter = "XML Configuration (*.xml)|*.xml|All Files (*.*)|*.*" # Set file type filter.

    $ofd.Title = "Select Infrastructure Configuration" # Set dialog title.

    

    if ($ofd.ShowDialog() -eq "OK") { # If user selects a file and clicks OK.

        Load-InfrastructureConfig -Path $ofd.FileName # Load configuration from the selected file.

    }

})



# 2. Save Log Button Action: Handles saving the content of the RichTextBox log to a file.

$btnSaveLog.Add_Click({

    $sfd = New-Object System.Windows.Forms.SaveFileDialog; # Create a new Save File Dialog.

    $sfd.Filter = "Text files (*.txt)|*.txt";             # Set file type filter.

    $sfd.FileName = "Decom_Log_$(Get-Date -f 'yyyyMMdd').txt" # Set default filename with current date.

    if($sfd.ShowDialog() -eq "OK") {                       # If user specifies a filename and clicks OK.

        ($rtbLog.Text) | Out-File $sfd.FileName;          # Write the RichTextBox content to the file.

        Log-Msg "Log saved to '$($sfd.FileName)'." "Cyan"   # Log confirmation.

    }

})



# 3. Login Button Action: Prompts the user for credentials to be used for subsequent operations.

$btnLogin.Add_Click({
    Log-Msg "Attempting login..." "White"
    # Audit: Start of login attempt
    Audit-Log -Action "Login Attempt" -Details "User attempting to log in." -Outcome "Info" -User "(Unknown)"

    if (-not $Script:AllowedADGroup) {
        Log-Msg "Error: Allowed AD Group is not configured in DecomConfig.xml. Please load a valid config." "Red"
        Audit-Log -Action "Login Attempt" -Details "Allowed AD Group not configured." -Outcome "Failure" -User "(System)"
        return
    }

    try { 
        $Cred = Get-Credential -Message "Enter credentials for access to Decommissioning Manager"
        if($Cred){ 
            $userName = $Cred.UserName
            
            Log-Msg "Validating user '$userName' against AD group '$Script:AllowedADGroup'..." "Cyan"
            Audit-Log -Action "Login Validation" -Target $Script:AllowedADGroup -Details "Validating user $userName" -Outcome "Info" -User $userName

            # Check if ActiveDirectory module is loaded for Test-ADGroupMembership
            if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
                Log-Msg "Error: ActiveDirectory module is not loaded. Cannot perform AD group membership check." "Red"
                [System.Windows.Forms.MessageBox]::Show("ActiveDirectory module is not loaded. Please ensure it's installed and available for import.", "Module Error", 0, 16)
                Audit-Log -Action "Login Validation" -Target $Script:AllowedADGroup -Details "ActiveDirectory module not loaded." -Outcome "Failure" -User $userName
                return
            }
            Import-Module ActiveDirectory -ErrorAction SilentlyContinue # Ensure it's imported in this scope

            if (Test-ADGroupMembership -UserName $userName -GroupName $Script:AllowedADGroup) {
                $Script:RunCreds = $Cred # Store credentials globally for other operations
                Log-Msg "Login successful for user '$userName'." "Lime"
                Audit-Log -Action "Login" -Target $Script:AllowedADGroup -Details "User $userName successfully logged in." -Outcome "Success" -User $userName
                # Enable all GUI controls
                $txtVM.Enabled = $true
                $txtTicket.Enabled = $true
                $chkPhys.Enabled = $true
                $chkCitrix.Enabled = $true
                $btnCheck.Enabled = $true
                # Other action buttons will be enabled after discovery
                # Update the dedicated label for logged-in user
                $lblLoggedInUser.Text = "Logged in as: $userName"
                $lblLoggedInUser.ForeColor = $Theme.AccentGreen
                
                # Update login button state
                $btnLogin.Text = "LOGGED IN" # Simpler text for the button
                $btnLogin.BackColor = $Theme.AccentGreen
                $btnLogin.Enabled = $false # Disable login button after successful login

            } else {
                $Script:RunCreds = $null # Clear any previously stored credentials
                Log-Msg "Login failed: User '$userName' is not a member of '$Script:AllowedADGroup'." "Red"
                Audit-Log -Action "Login" -Target $Script:AllowedADGroup -Details "User $userName failed group membership check." -Outcome "Failure" -User $userName
                # Keep controls disabled
                $txtVM.Enabled = $false
                $txtTicket.Enabled = $false
                $chkPhys.Enabled = $false
                $chkCitrix.Enabled = $false
                $btnCheck.Enabled = $false
                $btnStop.Enabled = $false
                $btnDecom.Enabled = $false
                $btnBackup.Enabled = $false
                $btnDel.Enabled = $false
            }
        } else {
            Log-Msg "Credential prompt cancelled by user or no credentials provided." "Yellow"
            Audit-Log -Action "Login Attempt" -Details "User cancelled credential prompt." -Outcome "Cancelled" -User "(Unknown)"
        }
    } catch { 
        Log-Msg "Login process failed: $($_.Exception.Message)" "Red"
        Audit-Log -Action "Login Attempt" -Details "Login process exception: $($_.Exception.Message)" -Outcome "Error" -User ($userName | default "(Unknown)")
    }
})

# Event handler for search button
$btnLogSearch.Add_Click({
    $searchTerm = $txtLogSearch.Text.Trim()
    
    # Clear previous highlights and reset colors
    $rtbLog.SelectAll()
    $rtbLog.SelectionBackColor = $rtbLog.BackColor # Reset background color
    $rtbLog.SelectionColor = $Theme.AccentGreen # Reset text color to default log color

    if ([string]::IsNullOrEmpty($searchTerm)) {
        # If search term is empty, just clear highlights and return
        return
    }

    # Find and highlight matches
    $startIndex = 0
    while ($startIndex -lt $rtbLog.TextLength) {
        $foundIndex = $rtbLog.Find($searchTerm, $startIndex, [System.Windows.Forms.RichTextBoxFindFlags]::None)
        if ($foundIndex -eq -1) {
            break # No more matches
        }
        
        $rtbLog.Select($foundIndex, $searchTerm.Length)
        $rtbLog.SelectionBackColor = [System.Drawing.Color]::Yellow # Highlight color
        $rtbLog.SelectionColor = [System.Drawing.Color]::Black # Make text readable on yellow
        
        $startIndex = $foundIndex + $searchTerm.Length
    }
    
    # Scroll to the first match if found
    $firstMatchIndex = $rtbLog.Find($searchTerm, 0, [System.Windows.Forms.RichTextBoxFindFlags]::None)
    if ($firstMatchIndex -ne -1) {
        $rtbLog.SelectionStart = $firstMatchIndex
        $rtbLog.ScrollToCaret()
    }
})

# --- DISCOVERY LOGIC (BATCH LOOP) ---
# This event handler performs an auto-discovery process for the entered target hostnames.
# It checks for machine presence in Active Directory and System Center Virtual Machine Manager (VMM).
$btnCheck.Add_Click({
    $Targets = Get-TargetList # Get the list of target hostnames from the input field.
    $currentUser = $Script:RunCreds.UserName # Get current user for auditing

    # Audit: Start of Discovery
    Audit-Log -Action "Discovery" -Target ($Targets -join ", ") -Outcome "Info" -Details "Starting discovery for $($Targets.Count) targets." -User $currentUser

    if ($Targets.Count -eq 0) { 
        Log-Msg "No target hostnames provided for discovery." "Yellow"
        Audit-Log -Action "Discovery" -Target "N/A" -Outcome "Warning" -Details "No target hostnames provided." -User $currentUser
        Update-ProgressLabel -Message "Idle - No targets for discovery" # NEW
        return 
    }

    $progressBar.Value = 0 # Reset progress bar at the start of discovery.
    Update-ProgressLabel -Message "Initializing Discovery..." # NEW
    $Script:IsPhysical = $chkPhys.Checked # Set global flag based on checkbox.
    $Script:VmmTargetHost = @{}             # Clear previous VMM host mappings.
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor # Change cursor to indicate busy state.
    
    # Reset State: Disable all action buttons and reset their colors.
    $btnStop.Enabled=$false; $btnDecom.Enabled=$false; $btnBackup.Enabled=$false; $btnDel.Enabled=$false
    $btnStop.BackColor=$Theme.BgControl; $btnDecom.BackColor=$Theme.BgControl; $btnBackup.BackColor=$Theme.BgControl; $btnDel.BackColor=$Theme.BgControl

    # NEW: Clear previous ListView items
    $lvTargets.Items.Clear()

    $Prefix = Get-TicketPrefix # Get log prefix from ticket number.
    $FoundCount = 0            # Counter for successfully discovered targets.
    $TotalTargets = $Targets.Count
    $CurrentTargetIndex = 0

    foreach ($Target in $Targets) {
        $CurrentTargetIndex++
        Update-ProgressLabel -Message "Discovering $Target ($CurrentTargetIndex of $TotalTargets)..." # NEW
        Log-Msg "${Prefix}Processing Discovery for: $Target" "White"; 
        [System.Windows.Forms.Application]::DoEvents() # Keep UI responsive.

        # NEW: Add target to ListView and store ListViewItem reference
        $lvi = New-Object System.Windows.Forms.ListViewItem($Target) # Hostname
        $lvi.SubItems.Add("Pending") # AD Status
        $lvi.SubItems.Add("Pending") # VMM Status
        $lvi.SubItems.Add("Discovering...") # Overall Status
        $lvi.SubItems.Add("Initializing Discovery") # Last Action
        $lvi.Tag = $Target # Store hostname in Tag for easy lookup
        $lvTargets.Items.Add($lvi) | Out-Null
        
        # 1. AD Discovery: Check if the computer object exists in the target OU.
        $lvi.SubItems[4].Text = "Checking AD..." # Update Last Action
        Update-ProgressLabel -Message "Discovering $Target (AD Check)..." # NEW
        try {
            $Searcher = [ADSISearcher]""
            $Searcher.SearchRoot = [ADSI]"LDAP://$Script:TargetOU" # Search within the configured OU.
            $Searcher.Filter = "(&(objectClass=computer)(name=$Target))" # Filter for computer objects by name.
            $Result = $Searcher.FindOne()
            if ($Result) {
                $DN = $Result.Properties.distinguishedname[0] # Get Distinguished Name.
                Log-Msg "${Prefix}  [AD] Verified: $DN" "Cyan"
                Audit-Log -Action "AD Discovery" -Target $Target -Outcome "Success" -Details "AD verified: $DN" -User $currentUser
                $lvi.SubItems[1].Text = "Found" # AD Status
                $lvi.SubItems[4].Text = "AD Verified" # Last Action
                $FoundCount++
            } else { 
                Log-Msg "${Prefix}  [AD] Not found in Target OU '$Script:TargetOU'." "Red" 
                Audit-Log -Action "AD Discovery" -Target $Target -Outcome "Failure" -Details "Not found in Target OU '$Script:TargetOU'." -User $currentUser
                $lvi.SubItems[1].Text = "Not Found" # AD Status
                $lvi.SubItems[4].Text = "AD Not Found" # Last Action
            }
        } catch { 
            $errorMessage = ($_.Exception).Message
            Log-Msg "${Prefix}  [AD] Error during AD lookup: $errorMessage" "Red" 
            Audit-Log -Action "AD Discovery" -Target $Target -Outcome "Error" -Details "AD lookup error: $errorMessage" -User $currentUser
            $lvi.SubItems[1].Text = "Error" # AD Status
            $lvi.SubItems[4].Text = "AD Error: $errorMessage" # Last Action
        }

        # 2. VMM Discovery (Skip if Physical): Check if it's a VM managed by SCVMM.
        if (!$Script:IsPhysical) {
            $lvi.SubItems[4].Text = "Checking VMM..." # Update Last Action
            Update-ProgressLabel -Message "Discovering $Target (VMM Check)..." # NEW
            $VmmFound = $false
            foreach ($Srv in $Script:VmmServers) {
                try {
                    $VM = Get-SCVirtualMachine -VMMServer $Srv -Name $Target -ErrorAction SilentlyContinue
                    if ($VM) {
                        Log-Msg "${Prefix}  [VMM] Found on host $($VM.VMHostName)." "Cyan"
                        Audit-Log -Action "VMM Discovery" -Target $Target -Outcome "Success" -Details "Found on host $($VM.VMHostName)." -User $currentUser
                        $lvi.SubItems[2].Text = "Found" # VMM Status
                        $lvi.SubItems[4].Text = "VMM Verified" # Last Action
                        $Script:VmmTargetHost[$Target] = $Srv # Map VM to its VMM server.
                        $VmmFound = $true; $FoundCount++
                        break # VM found, no need to check other VMM servers.
                    }
                } catch { 
                    $errorMessage = ($_.Exception).Message
                    Log-Msg "${Prefix}  [VMM] Error during VMM lookup on server ${Srv}: $errorMessage" "Red" 
                    Audit-Log -Action "VMM Discovery" -Target $Target -Outcome "Error" -Details ("VMM lookup error on " + $Srv + ": " + $errorMessage) -User $currentUser
                    $lvi.SubItems[2].Text = "Error" # VMM Status
                    $lvi.SubItems[4].Text = "VMM Error: $errorMessage" # Last Action
                }
            }
            if (!$VmmFound) { 
                Log-Msg "${Prefix}  [VMM] Not found on any configured VMM server." "Red" 
                Audit-Log -Action "VMM Discovery" -Target $Target -Outcome "Failure" -Details "Not found on any configured VMM server." -User $currentUser
                $lvi.SubItems[2].Text = "Not Found" # VMM Status
                $lvi.SubItems[4].Text = "VMM Not Found" # Last Action
            }
        } else {
            $lvi.SubItems[2].Text = "N/A (Physical)" # VMM Status
        }
        
        # Determine Overall Status after AD/VMM checks
        if ($lvi.SubItems[1].Text -eq "Found" -or $lvi.SubItems[2].Text -eq "Found" -or $lvi.SubItems[2].Text -eq "N/A (Physical)") {
            $lvi.SubItems[3].Text = "Ready for Action" # Overall Status
        } else {
            $lvi.SubItems[3].Text = "Not Found (Cannot Decom)" # Overall Status
        }

        $progress = [int](($CurrentTargetIndex / $TotalTargets) * 100)
        $progressBar.Value = $progress
        [System.Windows.Forms.Application]::DoEvents() # Update progress bar visually
    }

    # ENABLE BUTTONS IF MACHINES ARE FOUND: Activate action buttons if at least one target was discovered.
    if ($FoundCount -gt 0) {
        $btnStop.Enabled = $true; $btnDecom.Enabled = $true; $btnBackup.Enabled = $true # Enable main action buttons
        $btnStop.BackColor = $Theme.AccentWarn; $btnDecom.BackColor = $Theme.AccentGreen; $btnBackup.BackColor = $Theme.AccentBlue # Set their colors
        
        # PHYSICAL LOGIC: Adjust Delete VM button based on physical flag.
        if ($Script:IsPhysical) {
            $btnDel.Enabled = $false        # Disable Delete VM for physical hardware.
            $btnDel.BackColor = $Theme.BgControl
            Log-Msg "--- Discovery Complete (Physical Mode). 'Delete VM' is disabled. ---" "Yellow"
            Audit-Log -Action "Discovery Complete" -Target ($Targets -join ", ") -Outcome "Success" -Details "Discovery complete (Physical Mode), Delete VM disabled." -User $currentUser
            Update-ProgressLabel -Message "Discovery Complete (Physical Mode). Delete VM disabled."
        } else {
            $btnDel.Enabled = $true         # Enable Delete VM for virtual machines.
            $btnDel.BackColor = $Theme.AccentRed
            Log-Msg "--- Discovery Complete. Ready for Batch Actions. ---" "Lime"
            Audit-Log -Action "Discovery Complete" -Target ($Targets -join ", ") -Outcome "Success" -Details "Discovery complete, ready for batch actions." -User $currentUser
            Update-ProgressLabel -Message "Discovery Complete. Ready for batch actions."
        }
        
        # NEW: Send email after successful discovery
        $emailSubject = "Decommissioning Manager: Discovery Complete"
        $emailBody = "Discovery for the following targets has been completed successfully:`n`n"
        $emailBody += ($Targets | ForEach-Object { "- $_" }) -join "`n"
        $emailBody += "`n`nReady for batch actions. Current User: $currentUser"
        Send-DecomEmail -Subject $emailSubject -Body $emailBody

    } else {
        Log-Msg "--- No valid targets found. Check names or OU, or VMM connectivity. ---" "Orange"
        Audit-Log -Action "Discovery Complete" -Target ($Targets -join ", ") -Outcome "Failure" -Details "No valid targets found." -User $currentUser
        Update-ProgressLabel -Message "Discovery Complete - No valid targets found."
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default # Restore default cursor.
    $progressBar.Value = 100 # Ensure progress bar is full on completion
})

# --- STOP ACTION ---
# This event handler initiates the shutdown process for target machines.
# It differentiates between physical servers (using Stop-Computer) and VMM-managed VMs.
$btnStop.Add_Click({
    $Targets = Get-TargetList # Get the list of target hostnames.
    $currentUser = $Script:RunCreds.UserName # Get current user for auditing

    # Audit: Start of Stop Batch
    Audit-Log -Action "Stop Batch" -Target ($Targets -join ", ") -Outcome "Info" -Details "Starting Stop operation for $($Targets.Count) targets." -User $currentUser

    if ($Targets.Count -eq 0) { 
        Log-Msg "No targets to stop." "Yellow"
        Update-ProgressLabel -Message "Idle - No targets for Stop operation" # NEW
        return 
    }

    $Prefix = Get-TicketPrefix # Get log prefix.
    Log-Msg "${Prefix}--- STARTING BATCH SHUTDOWN ($($Targets.Count) Servers) ---" "White"
    Update-ProgressLabel -Message "Starting Stop operation for $($Targets.Count) targets..." # NEW
    $progressBar.Value = 0 # Reset progress bar at the start of shutdown.
    $TotalTargets = $Targets.Count
    $CurrentTargetIndex = 0
    
    $jobs = @()
    foreach ($Target in $Targets) {
        $CurrentTargetIndex++
        Update-ProgressLabel -Message "Stopping $Target ($CurrentTargetIndex of $TotalTargets)..." # NEW
        Log-Msg "${Prefix}[$Target] Starting Stop task as job..." "White"
        Audit-Log -Action "Stop Task Initiation" -Target $Target -Outcome "Info" -Details "Starting Stop task as job for $Target." -User $currentUser
        
        $job = Start-Job -FilePath ".\DecomWorker.ps1" `
                         -ArgumentList @(
                             "-Target", $Target,
                             "-Prefix", $Prefix,
                             "-IsPhysical", $Script:IsPhysical,
                             "-VmmTargetHost", $Script:VmmTargetHost
                         )
        $jobs += $job
        [System.Windows.Forms.Application]::DoEvents() # Keep UI responsive.
    }

    # Monitor and collect job results
    while ($jobs | Where-Object { $_.State -eq "Running" -or $_.State -eq "NotStarted" }) {
        $completedJobs = $jobs | Where-Object { $_.State -eq "Completed" -or $_.State -eq "Failed" }
        $progress = [int](($completedJobs.Count / $TotalTargets) * 100)
        $progressBar.Value = $progress
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200 # Avoid busy-waiting
    }

    foreach ($job in $jobs) {
        $result = Receive-Job -Job $job -Keep # Get results and keep job for inspection if needed
        $jobOutput = $result | Where-Object { $_ -is [PSCustomObject] -and $_.Phase -eq "Stop" } | Select-Object -First 1
        $workerLogs = $result | Where-Object { $_ -is [string] -and $_ -like "[WORKER]*" }
        
        foreach ($logLine in $workerLogs) {
            Log-Msg "$logLine" "Gray" # Display worker logs in main GUI
        }

        if ($jobOutput) {
            if ($jobOutput.Success) {
                Log-Msg "${Prefix}[$jobOutput.Target] Stop: Completed successfully." "Lime"
                Audit-Log -Action "Stop Task" -Target $jobOutput.Target -Outcome "Success" -Details "Stop operation completed." -User $currentUser
                Update-ListViewTargetStatus -TargetName $jobOutput.Target -OverallStatus "Stopped" -LastAction "Stop Completed" # NEW
            } else {
                Log-Msg "${Prefix}[$jobOutput.Target] Stop: FAILED! $($jobOutput.Message)" "Red"
                Audit-Log -Action "Stop Task" -Target $jobOutput.Target -Outcome "Failure" -Details "Stop operation failed: $($jobOutput.Message)" -User $currentUser
                Update-ListViewTargetStatus -TargetName $jobOutput.Target -OverallStatus "Stop Failed" -LastAction "Stop Failed: $($jobOutput.Message)" # NEW
            }
        } else {
            Log-Msg "${Prefix}[$job.Name] Stop: Job failed to return expected output. State: $($job.State)" "Red"
            Audit-Log -Action "Stop Task" -Target $job.Name -Outcome "Failure" -Details "Job failed to return expected output. State: $($job.State)" -User $currentUser
            Update-ListViewTargetStatus -TargetName $job.Name -OverallStatus "Stop Failed" -LastAction "Job Output Error: $($job.State)" # NEW
        }
        Remove-Job -Job $job # Clean up job
    }
    Log-Msg "${Prefix}--- BATCH SHUTDOWN COMPLETE ---" "White"
    Audit-Log -Action "Stop Batch" -Target ($Targets -join ", ") -Outcome "Success" -Details "Batch shutdown completed." -User $currentUser
    Update-ProgressLabel -Message "Stop Batch Complete. Processed $($TotalTargets) targets." # NEW
    $progressBar.Value = 100 # Ensure progress bar is full.
})

# --- LOGICAL CLEANUP (CITRIX SAFE + BATCH) ---
# This event handler performs various logical cleanup tasks across different infrastructure services.
# These tasks typically remove references to the decommissioned machine without affecting its storage.
$btnDecom.Add_Click({
    $Targets = Get-TargetList # Get the list of target hostnames.
    $currentUser = $Script:RunCreds.UserName # Get current user for auditing

    # Audit: Start of Logical Clean Batch
    Audit-Log -Action "Logical Clean Batch" -Target ($Targets -join ", ") -Outcome "Info" -Details "Starting Logical Clean operation for $($Targets.Count) targets." -User $currentUser

    if ($Targets.Count -eq 0) { 
        Log-Msg "No targets for logical cleanup." "Yellow"
        Update-ProgressLabel -Message "Idle - No targets for Logical Clean operation" # NEW
        return 
    }

    $Prefix = Get-TicketPrefix # Get log prefix.
    
    # User confirmation for this potentially destructive batch operation.
    $Confirm = [System.Windows.Forms.MessageBox]::Show("Confirm LOGICAL CLEANUP on $($Targets.Count) servers?", "Confirm Batch", 4, 32)
    if ($Confirm -ne "Yes") { 
        Log-Msg "${Prefix}ABORTED: Logical cleanup cancelled by user." "Yellow"
        Audit-Log -Action "Logical Clean Batch" -Target ($Targets -join ", ") -Outcome "Cancelled" -Details "Logical cleanup cancelled by user confirmation." -User $currentUser
        Update-ProgressLabel -Message "Logical Clean Operation Cancelled" # NEW
        return 
    } # Abort if user does not confirm.

    Log-Msg "${Prefix}--- STARTING BATCH LOGICAL CLEANUP ---" "White"
    Update-ProgressLabel -Message "Starting Logical Clean operation for $($Targets.Count) targets..." # NEW
    $progressBar.Value = 0 # Reset progress bar at the start of logical cleanup.
    $TotalTargets = $Targets.Count
    $CurrentTargetIndex = 0
    
    $jobs = @()
    foreach ($Target in $Targets) {
        $CurrentTargetIndex++
        Update-ProgressLabel -Message "Cleaning $Target ($CurrentTargetIndex of $TotalTargets)..." # NEW
        Log-Msg "${Prefix}[$Target] Starting Logical Clean task as job..." "White"
        Audit-Log -Action "Logical Clean Task Initiation" -Target $Target -Outcome "Info" -Details "Starting Logical Clean task as job for $Target." -User $currentUser
        
        $job = Start-Job -FilePath ".\DecomWorker.ps1" `
                         -ArgumentList @(
                             "-Target", $Target,
                             "-Prefix", $Prefix,
                             "-Phase", "Clean", # Indicate to worker to perform cleanup phases
                             "-IsPhysical", $Script:IsPhysical,
                             "-IncludeCitrixCleanup", $chkCitrix.Checked,
                             "-TargetOU", $Script:TargetOU,
                             "-SccmSiteCode", $Script:SccmSiteCode,
                             "-SccmProvider", $Script:SccmProvider,
                             "-DC", $Script:DC,
                             "-ForwardZone", $Script:ForwardZone,
                             "-VmmTargetHost", $Script:VmmTargetHost # Need VMM mapping for Citrix path
                         )
        $jobs += $job
        [System.Windows.Forms.Application]::DoEvents() # Keep UI responsive.
    }

    # Monitor and collect job results
    while ($jobs | Where-Object { $_.State -eq "Running" -or $_.State -eq "NotStarted" }) {
        $completedJobs = $jobs | Where-Object { $_.State -eq "Completed" -or $_.State -eq "Failed" }
        $progress = [int](($completedJobs.Count / $TotalTargets) * 100)
        $progressBar.Value = $progress
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200 # Avoid busy-waiting
    }

    foreach ($job in $jobs) {
        $result = Receive-Job -Job $job -Keep # Get results and keep job for inspection if needed
        $jobOutput = $result | Where-Object { $_ -is [PSCustomObject] -and $_.Phase -eq "Clean" } | Select-Object -First 1
        $workerLogs = $result | Where-Object { $_ -is [string] -and $_ -like "[WORKER]*" }
        
        foreach ($logLine in $workerLogs) {
            Log-Msg "$logLine" "Gray" # Display worker logs in main GUI
        }

        if ($jobOutput) {
            if ($jobOutput.Success) {
                Log-Msg "${Prefix}[$jobOutput.Target] Logical Clean: Completed successfully." "Lime"
                Audit-Log -Action "Logical Clean Task" -Target $jobOutput.Target -Outcome "Success" -Details "Logical Clean operation completed." -User $currentUser
                Update-ListViewTargetStatus -TargetName $jobOutput.Target -OverallStatus "Cleaned" -LastAction "Logical Clean Completed" # NEW
            } else {
                Log-Msg "${Prefix}[$jobOutput.Target] Logical Clean: FAILED! $($jobOutput.Message)" "Red"
                Audit-Log -Action "Logical Clean Task" -Target $jobOutput.Target -Outcome "Failure" -Details "Logical Clean operation failed: $($jobOutput.Message)" -User $currentUser
                Update-ListViewTargetStatus -TargetName $jobOutput.Target -OverallStatus "Clean Failed" -LastAction "Clean Failed: $($jobOutput.Message)" # NEW
            }
        } else {
            Log-Msg "${Prefix}[$job.Name] Logical Clean: Job failed to return expected output. State: $($job.State)" "Red"
            Audit-Log -Action "Logical Clean Task" -Target $job.Name -Outcome "Failure" -Details "Job failed to return expected output. State: $($job.State)" -User $currentUser
            Update-ListViewTargetStatus -TargetName $job.Name -OverallStatus "Clean Failed" -LastAction "Job Output Error: $($job.State)" # NEW
        }
        Remove-Job -Job $job # Clean up job
    }
    Log-Msg "${Prefix}--- BATCH LOGICAL CLEANUP COMPLETE ---" "White"
    Audit-Log -Action "Logical Clean Batch" -Target ($Targets -join ", ") -Outcome "Success" -Details "Batch logical cleanup completed." -User $currentUser
    Update-ProgressLabel -Message "Logical Clean Batch Complete. Processed $($TotalTargets) targets." # NEW
    $progressBar.Value = 100 # Ensure progress bar is full.
})

# --- DELETE FROM HYPERVISOR (BLOCKED IF PHYSICAL) ---
# This event handler initiates the permanent deletion of virtual machines from the hypervisor (VMM).
# It includes critical safety checks to prevent accidental deletion of physical hardware.
$btnDel.Add_Click({
    $Targets = Get-TargetList # Get the list of target hostnames.
    $currentUser = $Script:RunCreds.UserName # Get current user for auditing

    # Audit: Start of Delete VM Batch
    Audit-Log -Action "Delete VM Batch" -Target ($Targets -join ", ") -Outcome "Info" -Details "Starting Delete VM operation for $($Targets.Count) targets." -User $currentUser

    if ($Targets.Count -eq 0) { 
        Log-Msg "No targets for VM deletion." "Yellow"
        Update-ProgressLabel -Message "Idle - No targets for Delete VM operation" # NEW
        return 
    }
    
    $Prefix = Get-TicketPrefix # Get log prefix.
    
    # DOUBLE SAFETY CHECK: Prevent deletion if 'Physical Hardware' checkbox is ticked.
    if ($Script:IsPhysical) { 
        [System.Windows.Forms.MessageBox]::Show("Operation not allowed in Physical Mode. 'Delete VM' is disabled.", "Blocked Operation", 0, 16)
        Log-Msg "${Prefix}ABORTED: Physical hardware protection is active. Cannot delete from hypervisor." "Red"
        Audit-Log -Action "Delete VM Batch" -Target ($Targets -join ", ") -Outcome "Blocked" -Details "Physical hardware protection active. Cannot delete." -User $currentUser
        Update-ProgressLabel -Message "Delete VM Blocked (Physical Hardware)" # NEW
        return # Abort the operation.
    }
    
    # Critical confirmation: Requires user to type 'DELETE' to proceed with irreversible action.
    $InputUser = [Microsoft.VisualBasic.Interaction]::InputBox("DANGER: PERMANENTLY DELETING $($Targets.Count) VMs.`nType 'DELETE' to confirm:", "Confirm Destructive Action", "")
    
    if ($InputUser -ne "DELETE") { 
        Log-Msg "${Prefix}ABORTED: VM deletion cancelled by user." "Yellow"
        Audit-Log -Action "Delete VM Batch" -Target ($Targets -join ", ") -Outcome "Cancelled" -Details "VM deletion cancelled by user confirmation." -User $currentUser
        Update-ProgressLabel -Message "Delete VM Operation Cancelled" # NEW
        return # Abort if confirmation text is incorrect.
    }

    Log-Msg "${Prefix}--- STARTING BATCH VM DELETION ---" "Red"
    Update-ProgressLabel -Message "Starting Delete VM operation for $($Targets.Count) targets..." # NEW
    $progressBar.Value = 0 # Reset progress bar at the start of VM deletion.
    $TotalTargets = $Targets.Count
    $CurrentTargetIndex = 0
    
    $jobs = @()
    foreach ($Target in $Targets) {
        $CurrentTargetIndex++
        Update-ProgressLabel -Message "Deleting $Target ($CurrentTargetIndex of $TotalTargets)..." # NEW
        Log-Msg "${Prefix}[$Target] Starting Delete VM task as job..." "White"
        Audit-Log -Action "Delete VM Task Initiation" -Target $Target -Outcome "Info" -Details "Starting Delete VM task as job for $Target." -User $currentUser
        
        $job = Start-Job -FilePath ".\DecomWorker.ps1" `
                         -ArgumentList @(
                             "-Target", $Target,
                             "-Prefix", $Prefix,
                             "-Phase", "Delete", # Indicate to worker to perform delete phase
                             "-IsPhysical", $Script:IsPhysical,
                             "-VmmTargetHost", $Script:VmmTargetHost
                         )
        $jobs += $job
        [System.Windows.Forms.Application]::DoEvents() # Keep UI responsive.
    }

    # Monitor and collect job results
    while ($jobs | Where-Object { $_.State -eq "Running" -or $_.State -eq "NotStarted" }) {
        $completedJobs = $jobs | Where-Object { $_.State -eq "Completed" -or $_.State -eq "Failed" }
        $progress = [int](($completedJobs.Count / $TotalTargets) * 100)
        $progressBar.Value = $progress
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200 # Avoid busy-waiting
    }

    foreach ($job in $jobs) {
        $result = Receive-Job -Job $job -Keep # Get results and keep job for inspection if needed
        $jobOutput = $result | Where-Object { $_ -is [PSCustomObject] -and $_.Phase -eq "Delete" } | Select-Object -First 1
        $workerLogs = $result | Where-Object { $_ -is [string] -and $_ -like "[WORKER]*" }
        
        foreach ($logLine in $workerLogs) {
            Log-Msg "$logLine" "Gray" # Display worker logs in main GUI
        }

        if ($jobOutput) {
            if ($jobOutput.Success) {
                Log-Msg "${Prefix}[$jobOutput.Target] Delete VM: Completed successfully." "Lime"
                Audit-Log -Action "Delete VM Task" -Target $jobOutput.Target -Outcome "Success" -Details "Delete VM operation completed." -User $currentUser
                Update-ListViewTargetStatus -TargetName $jobOutput.Target -OverallStatus "DELETED" -LastAction "Delete Completed" # NEW
            } else {
                Log-Msg "${Prefix}[$jobOutput.Target] Delete VM: FAILED! $($jobOutput.Message)" "Red"
                Audit-Log -Action "Delete VM Task" -Target $jobOutput.Target -Outcome "Failure" -Details "Delete VM operation failed: $($jobOutput.Message)" -User $currentUser
                Update-ListViewTargetStatus -TargetName $jobOutput.Target -OverallStatus "Delete Failed" -LastAction "Delete Failed: $($jobOutput.Message)" # NEW
            }
        } else {
            Log-Msg "${Prefix}[$job.Name] Delete VM: Job failed to return expected output. State: $($job.State)" "Red"
            Audit-Log -Action "Delete VM Task" -Target $job.Name -Outcome "Failure" -Details "Job failed to return expected output. State: $($job.State)" -User $currentUser
            Update-ListViewTargetStatus -TargetName $job.Name -OverallStatus "Delete Failed" -LastAction "Job Output Error: $($job.State)" # NEW
        }
        Remove-Job -Job $job # Clean up job
    }
    Log-Msg "${Prefix}--- BATCH VM DELETION COMPLETE ---" "White"
    Audit-Log -Action "Delete VM Batch" -Target ($Targets -join ", ") -Outcome "Success" -Details "Batch VM deletion completed." -User $currentUser
    Update-ProgressLabel -Message "Delete VM Batch Complete. Processed $($TotalTargets) targets." # NEW
    $progressBar.Value = 100 # Ensure progress bar is full.
})

# --- AUTO LOAD CONFIG AT STARTUP ---
# Attempts to automatically load the DecomConfig.xml file located in the same directory as the script.
# This provides default settings upon launching the application.
if ([string]::IsNullOrEmpty($PSScriptRoot)) { # Check if $PSScriptRoot is defined (it is when run as a script).
    $BaseDir = (Get-Location).Path           # If not, use current working directory.
} else { 
    $BaseDir = $PSScriptRoot                 # Otherwise, use the script's directory.
}
$AutoXml = Join-Path $BaseDir "DecomConfig.xml" # Construct full path to the config file.
if (Test-Path $AutoXml) { 
    Load-InfrastructureConfig $AutoXml       # Attempt to load configuration.
} else { 
    Log-Msg "Info: DecomConfig.xml not found. Using defaults. Load manually if needed." "Yellow" # Inform user.
}

# Displays the main GUI form and starts the message loop.
# Out-Null prevents the form object from being written to the console.
# Dispose() releases system resources used by the form after it's closed.
$form.ShowDialog() | Out-Null; 

# NEW: Send session summary email upon application exit
$currentUser = $Script:RunCreds.UserName # Get current user for auditing
if (-not [string]::IsNullOrEmpty($Script:AuditLogPath) -and (Test-Path $Script:AuditLogPath)) {
    try {
        # Get session logs. Assuming the audit log is session-specific or we want recent entries.
        # This will get all new lines since last launch for a continuously appended log.
        # Alternatively, could filter by timestamp if the log is very large and persistent.
        $sessionLogs = Get-Content -Path $Script:AuditLogPath | Select-Object -Last 100 # Get last 100 lines as summary
        $emailSubject = "Decommissioning Manager: Session Summary for $currentUser ($(Get-Date -Format 'yyyy-MM-dd HH:mm'))"
        $emailBody = "Below is a summary of activities performed during the last Decommissioning Manager session:`n`n"
        $emailBody += ($sessionLogs -join "`n")
        $emailBody += "`n`nFor full details, please refer to the audit log file at: $Script:AuditLogPath"
        Send-DecomEmail -Subject $emailSubject -Body $emailBody

    } catch {
        Log-Msg "Error sending session summary email: $($_.Exception.Message)" "Red"
        # Audit-Log cannot be called here as application is closing, and it might depend on the log file itself.
    }
} else {
    Log-Msg "Audit log path not configured or file not found. Skipping session summary email." "Yellow"
}

$form.Dispose()
