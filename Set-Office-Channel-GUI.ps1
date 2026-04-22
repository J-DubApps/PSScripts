#Requires -Version 5.1
<#
.SYNOPSIS
    Office Update Channel Configuration GUI -- Microsoft 365 Apps.

.DESCRIPTION
    Sets the Microsoft 365 Apps update channel via the HKLM update policy and
    optionally constrains/locks the in-app channel selector exposure.

    Channels available (fastest -> slowest):
        - Beta Channel              (Fast Ring; Insider Fast)         [UNSUPPORTED in production]
        - Current Channel (Preview) (Slow Ring; ex Monthly Targeted)
        - Current Channel
        - Monthly Enterprise Channel
        - Semi-Annual Enterprise Channel
        - (Optional) Semi-Annual Enterprise Channel (Preview) -- see $Channels block

    Modern Microsoft 'updatebranch' strings are used for writes. Legacy aliases
    (InsiderFast, FirstReleaseCurrent, SemiAnnual, SemiAnnualPreview) are still
    recognized when reading existing registry state.

    Documented references (verified 2026-04):
      https://learn.microsoft.com/en-us/microsoft-365-apps/insider/deploy/registry
      https://learn.microsoft.com/en-us/microsoft-365-apps/updates/change-update-channels
      https://learn.microsoft.com/en-us/intune/device-configuration/settings-catalog/update-office

.NOTES
    Author    : Julian West (revised by Claude, 2026-04-22)
    PS edition: 5.1 target (splatting; no backtick continuations)
    Requires  : Local Administrator (writes HKLM)

    Does NOT touch HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration
    (CDNBaseUrl / UpdateChannel). Direct manipulation of those values is
    unsupported. The supported flow is: set 'updatebranch' here, then let the
    'Office Automatic Updates 2.0' scheduled task pick it up.
#>

# ============================================================================
# Assemblies & elevation
# ============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    [void][System.Windows.Forms.MessageBox]::Show(
        'Please run this script as Administrator.',
        'Elevation Required',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    exit 1
}

# ============================================================================
# Registry constants
# ============================================================================
$RegPath_OfficeUpdate    = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate'
$RegPath_ChannelExposure = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officechannelexposure'
$RegPath_OfficeCommon    = 'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common'
$ValueName_UpdateBranch  = 'updatebranch'
$ValueName_Suppress      = 'suppressofficechannelselector'

# ============================================================================
# Logging
# ============================================================================
$LogDir  = Join-Path $env:ProgramData 'OfficeChannelGUI'
$LogPath = Join-Path $LogDir 'OfficeChannelGUI.log'
if (-not (Test-Path $LogDir)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ValidateSet('INFO','WARN','ERROR')] [string] $Level = 'INFO'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = '{0}  [{1,-5}]  {2}\{3}  {4}' -f $ts, $Level, $env:USERDOMAIN, $env:USERNAME, $Message
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
}

# ============================================================================
# Channel catalog -- ordered fastest -> slowest
#   Key   = the 'updatebranch' REG_SZ value to write (MODERN Microsoft naming)
#   Value = friendly display string for the dropdown
# ============================================================================
$Channels = [ordered]@{
    'BetaChannel'          = 'Beta Channel  (Fast Ring / Insider Fast)  -- UNSUPPORTED for production'
    'CurrentPreview'       = 'Current Channel (Preview)  (Slow Ring / formerly Monthly Targeted)'
    'Current'              = 'Current Channel'
    'MonthlyEnterprise'    = 'Monthly Enterprise Channel'
    # Uncomment to expose Semi-Annual Enterprise Channel (Preview):
    # 'FirstReleaseDeferred' = 'Semi-Annual Enterprise Channel (Preview)'
    'Deferred'             = 'Semi-Annual Enterprise Channel'
}

# Legacy 'updatebranch' values -> modern equivalents (read-side compat only)
$LegacyAliases = @{
    'InsiderFast'         = 'BetaChannel'
    'FirstReleaseCurrent' = 'CurrentPreview'
    'SemiAnnual'          = 'Deferred'
    'SemiAnnualPreview'   = 'FirstReleaseDeferred'
}

# officechannelexposure DWord names -- these still use the OLDER channel naming
$ExposureKeyMap = @{
    'BetaChannel'          = 'insiderfast'
    'CurrentPreview'       = 'firstreleasecurrent'
    'Current'              = 'current'
    'MonthlyEnterprise'    = 'monthlyenterprise'
    'FirstReleaseDeferred' = 'firstreleasedeferred'
    'Deferred'             = 'deferred'
}

# Channels that warrant a "non-production" confirmation prompt
$NonProductionChannels = @('BetaChannel','CurrentPreview','FirstReleaseDeferred')

# ============================================================================
# Read current registry state (translates legacy aliases)
# ============================================================================
$currentRaw = $null
try {
    $currentRaw = (Get-ItemProperty -Path $RegPath_OfficeUpdate -Name $ValueName_UpdateBranch -ErrorAction Stop).$ValueName_UpdateBranch
} catch {
    $currentRaw = $null
}

$currentKey = $currentRaw
if ($currentRaw -and $LegacyAliases.ContainsKey($currentRaw)) {
    $currentKey = $LegacyAliases[$currentRaw]
}

if ($currentKey -and $Channels.Contains($currentKey)) {
    $currentFriendly = $Channels[$currentKey]
} elseif ($currentRaw) {
    $currentFriendly = "<unrecognized value: '$currentRaw'>"
} else {
    $currentFriendly = '<not set>'
}

Write-Log "GUI launched. Current registry value: '$currentRaw' -> '$currentFriendly'"

# ============================================================================
# Helper functions
# ============================================================================
function Get-ChannelKey {
    param([string] $FriendlyValue)
    foreach ($entry in $Channels.GetEnumerator()) {
        if ($entry.Value -eq $FriendlyValue) { return $entry.Key }
    }
    return $null
}

# ============================================================================
# Build the form
# ============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text            = 'Microsoft 365 Apps -- Update Channel Configuration'
$form.Size            = New-Object System.Drawing.Size(640, 380)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false

# --- Header ---
$lblHeader = New-Object System.Windows.Forms.Label
$lblHeader.AutoSize = $true
$lblHeader.Location = New-Object System.Drawing.Point(12, 12)
$lblHeader.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$lblHeader.Text     = 'Select desired Microsoft 365 Apps update channel:'
$form.Controls.Add($lblHeader)

# --- Currently-set indicator ---
$lblCurrent = New-Object System.Windows.Forms.Label
$lblCurrent.AutoSize  = $true
$lblCurrent.Location  = New-Object System.Drawing.Point(12, 36)
$lblCurrent.ForeColor = [System.Drawing.Color]::DimGray
$lblCurrent.Text      = "Currently set in registry:  $currentFriendly"
$form.Controls.Add($lblCurrent)

# --- Channel dropdown ---
$combo = New-Object System.Windows.Forms.ComboBox
$combo.DropDownStyle = 'DropDownList'
$combo.Width         = 600
$combo.Location      = New-Object System.Drawing.Point(12, 64)
$combo.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
foreach ($entry in $Channels.GetEnumerator()) {
    [void]$combo.Items.Add($entry.Value)
}
if ($currentKey -and $Channels.Contains($currentKey)) {
    $combo.SelectedItem = $Channels[$currentKey]
} else {
    # Default to Current Channel if nothing valid is set
    $combo.SelectedItem = $Channels['Current']
}
$form.Controls.Add($combo)

# --- Warning / status pane (color changes with selection) ---
$lblWarning = New-Object System.Windows.Forms.Label
$lblWarning.Location = New-Object System.Drawing.Point(12, 100)
$lblWarning.Size     = New-Object System.Drawing.Size(600, 70)
$lblWarning.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
$form.Controls.Add($lblWarning)

# --- Option: lock UI to selected channel only ---
$chkLockUI = New-Object System.Windows.Forms.CheckBox
$chkLockUI.AutoSize = $true
$chkLockUI.Location = New-Object System.Drawing.Point(12, 188)
$chkLockUI.Text     = 'Lock Office UI to ONLY the selected channel (hide all other choices)'
$chkLockUI.Checked  = $false
$form.Controls.Add($chkLockUI)

# --- Option: show channel selector in Office UI ---
$chkShowSelector = New-Object System.Windows.Forms.CheckBox
$chkShowSelector.AutoSize = $true
$chkShowSelector.Location = New-Object System.Drawing.Point(12, 214)
$chkShowSelector.Text     = 'Show "Update Channel" selector in Office (Word/Excel) > File > Account'
$chkShowSelector.Checked  = $true
$form.Controls.Add($chkShowSelector)

# --- Live warning updater ---
function Update-WarningLabel {
    $sel = Get-ChannelKey -FriendlyValue $combo.SelectedItem
    if (-not $sel) {
        $lblWarning.Text = ''
        return
    }
    switch ($sel) {
        'BetaChannel' {
            $lblWarning.ForeColor = [System.Drawing.Color]::Firebrick
            $lblWarning.Text = "WARNING: Beta Channel (Fast Ring) is NOT SUPPORTED by Microsoft for production use." + [Environment]::NewLine + "Frequent builds may include broken or unfinished features. Use only on lab / opt-in test machines."
        }
        'CurrentPreview' {
            $lblWarning.ForeColor = [System.Drawing.Color]::DarkOrange
            $lblWarning.Text = "CAUTION: Current Channel (Preview) -- the 'Slow Ring'. Receives builds ~1 week before Current Channel." + [Environment]::NewLine + "Intended for IT validation rings, not full-fleet production."
        }
        'FirstReleaseDeferred' {
            $lblWarning.ForeColor = [System.Drawing.Color]::DarkOrange
            $lblWarning.Text = "CAUTION: Semi-Annual Enterprise Channel (Preview). Validation ring for SAEC." + [Environment]::NewLine + "Typically deployed to ~5% of fleet for SAEC pre-release validation."
        }
        Default {
            $lblWarning.ForeColor = [System.Drawing.Color]::DarkGreen
            $lblWarning.Text = 'Production-supported channel.'
        }
    }
}
$combo.Add_SelectedIndexChanged({ Update-WarningLabel })
Update-WarningLabel  # initial paint

# --- Apply button ---
$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text     = 'Apply'
$btnSave.Width    = 95
$btnSave.Location = New-Object System.Drawing.Point(420, 295)
$btnSave.Add_Click({
    $selKey = Get-ChannelKey -FriendlyValue $combo.SelectedItem
    if (-not $selKey) {
        [void][System.Windows.Forms.MessageBox]::Show(
            'No valid channel selected.', 'Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Confirm if non-production channel
    if ($NonProductionChannels -contains $selKey) {
        $confirmText = "You are about to set this device to a non-production channel:" + [Environment]::NewLine + [Environment]::NewLine + "    $($combo.SelectedItem)" + [Environment]::NewLine + [Environment]::NewLine + "Proceed?"
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            $confirmText,
            'Confirm non-production channel',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "User cancelled non-production change to '$selKey'." -Level WARN
            return
        }
    }

    try {
        # 1) updatebranch (modern Microsoft string)
        if (-not (Test-Path $RegPath_OfficeUpdate)) {
            New-Item -Path $RegPath_OfficeUpdate -Force | Out-Null
        }
        Set-ItemProperty -Path $RegPath_OfficeUpdate -Name $ValueName_UpdateBranch -Value $selKey -Force

        # 2) officechannelexposure
        if (-not (Test-Path $RegPath_ChannelExposure)) {
            New-Item -Path $RegPath_ChannelExposure -Force | Out-Null
        }
        foreach ($ch in $ExposureKeyMap.GetEnumerator()) {
            $expose = 0
            if ($chkLockUI.Checked) {
                # Lock-down mode: expose only the chosen channel
                if ($ch.Key -eq $selKey) { $expose = 1 }
            } else {
                # Default: expose every channel offered in this script's dropdown
                if ($Channels.Contains($ch.Key)) { $expose = 1 }
            }
            $expoParams = @{
                Path         = $RegPath_ChannelExposure
                Name         = $ch.Value
                PropertyType = 'DWord'
                Value        = $expose
                Force        = $true
                ErrorAction  = 'Stop'
            }
            New-ItemProperty @expoParams | Out-Null
        }

        # 3) suppressofficechannelselector at parent 'common' key
        if (-not (Test-Path $RegPath_OfficeCommon)) {
            New-Item -Path $RegPath_OfficeCommon -Force | Out-Null
        }
        $suppressValue = if ($chkShowSelector.Checked) { 0 } else { 1 }
        $suppressParams = @{
            Path         = $RegPath_OfficeCommon
            Name         = $ValueName_Suppress
            PropertyType = 'DWord'
            Value        = $suppressValue
            Force        = $true
            ErrorAction  = 'Stop'
        }
        New-ItemProperty @suppressParams | Out-Null

        $lockMode = if ($chkLockUI.Checked) { 'lock-to-selected' } else { 'expose-all-listed' }
        $selMode  = if ($chkShowSelector.Checked) { 'selector-shown' } else { 'selector-suppressed' }
        Write-Log "Applied: updatebranch='$selKey'  exposure='$lockMode'  ui='$selMode'"

        $msg  = 'Update channel set to:' + [Environment]::NewLine
        $msg += "    $($combo.SelectedItem)" + [Environment]::NewLine + [Environment]::NewLine
        $msg += "Registry value 'updatebranch' = '$selKey'" + [Environment]::NewLine + [Environment]::NewLine
        $msg += 'Office will switch the channel within ~24h, or sooner if the' + [Environment]::NewLine
        $msg += "'Office Automatic Updates 2.0' scheduled task runs and an" + [Environment]::NewLine
        $msg += 'Office app then performs File > Account > Update Now.'
        [void][System.Windows.Forms.MessageBox]::Show(
            $msg, 'Success',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)

        # Refresh "currently set" line on the form
        $lblCurrent.Text = "Currently set in registry:  $($combo.SelectedItem)"
    } catch {
        Write-Log "FAILED: $($_.Exception.Message)" -Level ERROR
        [void][System.Windows.Forms.MessageBox]::Show(
            "Failed to write registry: $($_.Exception.Message)",
            'Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$form.Controls.Add($btnSave)

# --- Close button ---
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text     = 'Close'
$btnCancel.Width    = 95
$btnCancel.Location = New-Object System.Drawing.Point(523, 295)
$btnCancel.Add_Click({ $form.Close() })
$form.Controls.Add($btnCancel)

# ============================================================================
# Run
# ============================================================================
[System.Windows.Forms.Application]::EnableVisualStyles()
[void][System.Windows.Forms.Application]::Run($form)