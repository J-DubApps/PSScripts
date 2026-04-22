<#
.SYNOPSIS
    Configures the Microsoft Office update channel via registry policy (no UI).

.DESCRIPTION
    Sets the Office Update Channel policy for Office 16.0 by writing to:
        HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate

    Registry values applied (channel branch + automatic-update enablement):
      - updatebranch           : Channel code name (see -UpdateBranch)
      - enableautomaticupdates : DWORD 1
      - updateenabled          : DWORD 1
      - disableupdates         : DWORD 0

    Optionally also configures the in-app Office Update Channel selector
    visibility (the "Update Channel" button shown at File > Account >
    Update Options inside Word/Excel/etc.) via -OfficeUIChannelSelector.
    See that parameter's help for full behavior.

    Channel code-name mapping:
      InsiderFast       = Beta Channel
      InsiderSlow       = Current Channel (Preview)
      Current           = Current Channel
      MonthlyEnterprise = Monthly Enterprise Channel
      BroadPreview      = Semi-Annual Enterprise Channel (Preview)
      Broad             = Semi-Annual Enterprise Channel

    NOTE (July 2026):  Microsoft is converging Semi-Annual Enterprise Channel
    with Monthly Enterprise Channel's cadence.  Post-July 2026, SAC receives
    monthly feature updates and the effective support window shrinks to ~1 month
    (+ 2-month rollback).  MEC becomes the preferred enterprise channel.
    This script will emit a warning when selecting Broad or BroadPreview.

.PARAMETER UpdateBranch
    The code name of the channel to apply. Valid values:
      InsiderSlow, InsiderFast, Current, MonthlyEnterprise, BroadPreview, Broad

.PARAMETER OfficeUIChannelSelector
    Controls visibility of the in-app channel selector that end users see at
    File > Account > Update Options > Update Channel inside Office apps.

      Show     - Reveal the selector (suppressofficechannelselector = 0) AND
                 expose all six known channels via officechannelexposure so
                 users can see what they're on (or initiate a channel switch
                 themselves). Useful when IT is going to instruct users to
                 verify a channel change occasionally.

      Hide     - Suppress the selector (suppressofficechannelselector = 1).
                 Does NOT clear officechannelexposure values, so any prior
                 IT-controlled exposure lockdown is preserved.

      NoChange - (Default) Do not modify any UI selector or channel-exposure
                 registry values. Use this for ordinary unattended channel
                 switches where you don't want to disrupt existing UI policy.

.PARAMETER Force
    Suppress the interactive SAC deprecation warning prompt
    (for unattended / NinjaOne use).

.EXAMPLE
    # Interactive: Set to Monthly Enterprise Channel
    .\Set-OfficeUpdateChannel.ps1 -UpdateBranch MonthlyEnterprise

.EXAMPLE
    # NinjaOne (SYSTEM, unattended): switch to Current Channel, no UI changes
    powershell.exe -ExecutionPolicy Bypass -File "Set-OfficeUpdateChannel.ps1" -UpdateBranch Current -Force

.EXAMPLE
    # Switch to Beta Channel AND reveal the in-app selector so the user
    # can see they've moved (or self-toggle later)
    .\Set-OfficeUpdateChannel.ps1 -UpdateBranch InsiderFast -OfficeUIChannelSelector Show -Force

.EXAMPLE
    # NinjaOne: enforce MEC and lock down the UI (no end-user channel switching)
    powershell.exe -ExecutionPolicy Bypass -File "Set-OfficeUpdateChannel.ps1" -UpdateBranch MonthlyEnterprise -OfficeUIChannelSelector Hide -Force

.NOTES
    Requires Administrator privileges.
    Targets PowerShell 5.1 -- no backtick line-continuations used.
    Update channel is orthogonal to the M365 Copilot Release Audience
    (Frontier / Standard / Deferred), which is managed tenant-side in
    the M365 Admin Center, not via endpoint registry.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [ValidateSet(
        "InsiderSlow",
        "InsiderFast",
        "Current",
        "MonthlyEnterprise",
        "BroadPreview",
        "Broad"
    )]
    [string]$UpdateBranch,

    [Parameter()]
    [ValidateSet("Show", "Hide", "NoChange")]
    [string]$OfficeUIChannelSelector = "NoChange",

    [Parameter()]
    [switch]$Force
)

# -- Friendly name mapping (for logging) -------------------------------------
$ChannelDisplayNames = @{
    "InsiderFast"       = "Beta Channel"
    "InsiderSlow"       = "Current Channel (Preview)"
    "Current"           = "Current Channel"
    "MonthlyEnterprise" = "Monthly Enterprise Channel"
    "BroadPreview"      = "Semi-Annual Enterprise Channel (Preview)"
    "Broad"             = "Semi-Annual Enterprise Channel"
}

# -- officechannelexposure DWord names (UI exposure flags) -------------------
# These keys use the LEGACY Office channel naming and are independent of the
# value written to 'updatebranch'. They control which channels appear in the
# Office in-app channel selector when that selector is visible.
$ExposureKeyMap = @{
    "InsiderFast"       = "insiderfast"
    "InsiderSlow"       = "firstreleasecurrent"
    "Current"           = "current"
    "MonthlyEnterprise" = "monthlyenterprise"
    "BroadPreview"      = "firstreleasedeferred"
    "Broad"             = "deferred"
}

# -- Elevation check ---------------------------------------------------------
function Test-IsElevated {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsElevated)) {
    Write-Error "This script must be run with Administrator privileges."
    exit 1
}

# -- SAC deprecation / convergence warning -----------------------------------
$sacBranches = @("Broad", "BroadPreview")
if ($UpdateBranch -in $sacBranches) {
    $warnMsg = @(
        "WARNING: Semi-Annual Enterprise Channel is converging with Monthly"
        "Enterprise Channel in July 2026. Post-convergence, SAC receives monthly"
        "feature updates with a reduced support window (~1 month + 2-month rollback)."
        "Microsoft recommends migrating to MonthlyEnterprise for enterprise workloads."
        "Proceeding with '$UpdateBranch' as requested."
    ) -join " "

    Write-Warning $warnMsg

    if (-not $Force) {
        $response = Read-Host "Continue setting channel to '$UpdateBranch'? (Y/N)"
        if ($response -notin @("Y", "y", "Yes", "yes")) {
            Write-Host "Aborted by user."
            exit 0
        }
    }
}

# -- Registry configuration --------------------------------------------------
$regPath   = "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate"
$valueName = "updatebranch"

try {
    # Capture previous value for change logging
    $previousBranch = $null
    if (Test-Path $regPath) {
        $existingVal = Get-ItemProperty -Path $regPath -Name $valueName -ErrorAction SilentlyContinue
        if ($existingVal) {
            $previousBranch = $existingVal.$valueName
        }
    }

    # Ensure the policy key exists
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    # Set the branch value
    $setParams = @{
        Path        = $regPath
        Name        = $valueName
        Value       = $UpdateBranch
        ErrorAction = "Stop"
    }
    Set-ItemProperty @setParams

    # Ensure automatic updates are enabled
    $updateProps = @{
        "enableautomaticupdates" = 1
        "updateenabled"          = 1
        "disableupdates"         = 0
    }

    foreach ($propName in $updateProps.Keys) {
        $newItemParams = @{
            Path         = $regPath
            Name         = $propName
            PropertyType = "DWord"
            Value        = $updateProps[$propName]
            Force        = $true
        }
        New-ItemProperty @newItemParams | Out-Null
    }

    # -- Verification read-back (channel) ------------------------------------
    $verifyVal = (Get-ItemProperty -Path $regPath -Name $valueName).$valueName
    if ($verifyVal -ne $UpdateBranch) {
        Write-Error "Verification failed: registry shows '$verifyVal', expected '$UpdateBranch'."
        exit 1
    }

    # -- Office UI Channel Selector (optional) -------------------------------
    $uiChangeNote = ""
    if ($OfficeUIChannelSelector -ne "NoChange") {
        $commonPath        = "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common"
        $exposurePath      = "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officechannelexposure"
        $suppressValueName = "suppressofficechannelselector"

        if (-not (Test-Path $commonPath)) {
            New-Item -Path $commonPath -Force | Out-Null
        }

        switch ($OfficeUIChannelSelector) {

            "Show" {
                # Reveal the in-app selector
                $suppressShowParams = @{
                    Path         = $commonPath
                    Name         = $suppressValueName
                    PropertyType = "DWord"
                    Value        = 0
                    Force        = $true
                    ErrorAction  = "Stop"
                }
                New-ItemProperty @suppressShowParams | Out-Null

                # Expose all six known channels so the dropdown isn't empty
                if (-not (Test-Path $exposurePath)) {
                    New-Item -Path $exposurePath -Force | Out-Null
                }
                foreach ($expKeyName in $ExposureKeyMap.Values) {
                    $expParams = @{
                        Path         = $exposurePath
                        Name         = $expKeyName
                        PropertyType = "DWord"
                        Value        = 1
                        Force        = $true
                        ErrorAction  = "Stop"
                    }
                    New-ItemProperty @expParams | Out-Null
                }

                $uiChangeNote = " UI selector: SHOWN; all six channels exposed."
            }

            "Hide" {
                # Suppress the selector; leave officechannelexposure untouched
                $suppressHideParams = @{
                    Path         = $commonPath
                    Name         = $suppressValueName
                    PropertyType = "DWord"
                    Value        = 1
                    Force        = $true
                    ErrorAction  = "Stop"
                }
                New-ItemProperty @suppressHideParams | Out-Null

                $uiChangeNote = " UI selector: HIDDEN."
            }
        }

        # Verification read-back (UI selector). Non-fatal if it can't read.
        try {
            $verifySuppress = (Get-ItemProperty -Path $commonPath -Name $suppressValueName -ErrorAction Stop).$suppressValueName
            $expectedSuppress = if ($OfficeUIChannelSelector -eq "Show") { 0 } else { 1 }
            if ($verifySuppress -ne $expectedSuppress) {
                Write-Warning "UI selector verification: expected '$expectedSuppress', registry shows '$verifySuppress'."
            }
        } catch {
            Write-Warning "Unable to verify UI selector registry value: $_"
        }
    }

    # -- Success output ------------------------------------------------------
    $displayName = $ChannelDisplayNames[$UpdateBranch]
    $changeNote  = ""
    if ($previousBranch -and $previousBranch -ne $UpdateBranch) {
        $prevDisplay = $ChannelDisplayNames[$previousBranch]
        if (-not $prevDisplay) { $prevDisplay = $previousBranch }
        $changeNote = " (changed from '$prevDisplay')"
    }

    $successMsg = "Office update channel set to '$displayName' ($UpdateBranch)$changeNote. Update policies applied.$uiChangeNote"
    Write-Host $successMsg

    # Write to Application event log for audit trail
    $logParams = @{
        LogName   = "Application"
        Source    = "MSIInstaller"
        EventId   = 9100
        EntryType = "Information"
        Message   = $successMsg
    }
    try { Write-EventLog @logParams } catch { <# non-fatal #> }

    exit 0
}
catch {
    $errMsg = "Failed to configure Office Update Channel: $_"
    Write-Error $errMsg

    # Attempt to log failure to event log
    $errLogParams = @{
        LogName   = "Application"
        Source    = "MSIInstaller"
        EventId   = 9101
        EntryType = "Error"
        Message   = $errMsg
    }
    try { Write-EventLog @errLogParams } catch { <# non-fatal #> }

    exit 1
}