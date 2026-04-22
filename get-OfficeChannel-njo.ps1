<#
.SYNOPSIS
    Reports the Microsoft 365 Apps (Office 365) update channel for NinjaOne,
    with optional compliance check against an approved channel list.

.DESCRIPTION
    Silent NinjaOne-compatible reporter. No user-visible output; all output
    is written to NinjaOne's activity log as KEY=VALUE lines.

    Reports three channel signals so in-flight migrations and policy
    conflicts are visible, not hidden:
      - ActiveChannel : HKLM\...ClickToRun\Configuration!UpdateChannel
      - CDNChannel    : HKLM\...ClickToRun\Configuration!CDNBaseUrl
      - PolicyChannel : HKLM\...Policies\...\officeupdate (GPO or Intune)

    Plus version, architecture, installed product IDs, culture, and
    SharedComputerLicensing state.

    Performs a compliance check: ActiveChannel vs an approved list (defaults
    to Monthly Enterprise Channel and Semi-Annual Enterprise Channel, matching
    a typical law-firm / professional-services policy).

    Optionally writes ActiveChannel to a NinjaOne Custom Field.

.PARAMETER CustomField
    Optional NinjaOne Custom Field name to populate with the ActiveChannel
    friendly name. Leave empty to skip (default).

.PARAMETER AllowedChannels
    Array of friendly channel names considered compliant for this environment.
    Defaults to Monthly Enterprise Channel and Semi-Annual Enterprise Channel.

.PARAMETER FailOnNonCompliance
    When set, the script exits with code 2 if ActiveChannel is not in the
    AllowedChannels list. Use this pattern when you want a NinjaOne condition
    or automated remediation to trigger on drift. When not set (default),
    the script always exits 0 on successful execution and you filter the
    activity log / Custom Field for Compliant=False.

.EXAMPLE
    # Pattern A - silent inventory reporter (default)
    .\Get-OfficeUpdateChannel-NinjaOne.ps1

.EXAMPLE
    # Pattern A + populate a NinjaOne custom field
    .\Get-OfficeUpdateChannel-NinjaOne.ps1 -CustomField 'officeUpdateChannel'

.EXAMPLE
    # Pattern B - compliance trigger; non-approved channels exit 2
    .\Get-OfficeUpdateChannel-NinjaOne.ps1 -FailOnNonCompliance

.EXAMPLE
    # Custom approved list for a team that also permits Current Channel
    .\Get-OfficeUpdateChannel-NinjaOne.ps1 -AllowedChannels 'Monthly Enterprise Channel','Current Channel'

.NOTES
    Target:      PowerShell 5.1 (Windows PowerShell)
    Runs as:     SYSTEM under NinjaOne agent
    Exit codes:  0 = Script ran to completion (inspect RESULT= / Compliant=)
                 1 = Runtime error (see ErrorMessage=)
                 2 = Non-compliant ActiveChannel (only when -FailOnNonCompliance set)
#>

[CmdletBinding()]
param(
    [string]$CustomField = '',

    [string[]]$AllowedChannels = @(
        'Monthly Enterprise Channel',
        'Semi-Annual Enterprise Channel'
    ),

    [switch]$FailOnNonCompliance
)

# -------------------------------------------------------------------
# Channel GUID -> Friendly name mapping
# Verified against Microsoft Learn (Intune settings catalog and the
# Microsoft 365 Apps best-practices docs) as of April 2026.
# Add new channel GUIDs here as Microsoft introduces them.
# -------------------------------------------------------------------
$ChannelMap = @{
    '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = 'Current Channel'
    '64256afe-f5d9-4f86-8936-8840a6a4f5be' = 'Current Channel (Preview)'
    '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = 'Monthly Enterprise Channel'
    '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = 'Semi-Annual Enterprise Channel'
    'b8f9b850-328d-4355-9145-c59439a0c4cf' = 'Semi-Annual Enterprise Channel (Preview) [Deprecated]'
    '5440fd1f-7ecb-4221-8110-145efaa6372f' = 'Beta Channel'
    'f2e724c1-748f-4b47-8fb8-8e0d210e9208' = 'LTSC / Office 2019 / Perpetual VL'
}

# GPO sometimes writes a literal name instead of a URL
$PolicyNameMap = @{
    'Current'              = 'Current Channel'
    'CurrentPreview'       = 'Current Channel (Preview)'
    'FirstReleaseCurrent'  = 'Current Channel (Preview)'                  # legacy
    'MonthlyEnterprise'    = 'Monthly Enterprise Channel'
    'SemiAnnual'           = 'Semi-Annual Enterprise Channel'
    'Deferred'             = 'Semi-Annual Enterprise Channel'             # legacy
    'SemiAnnualPreview'    = 'Semi-Annual Enterprise Channel (Preview)'
    'FirstReleaseDeferred' = 'Semi-Annual Enterprise Channel (Preview)'   # legacy
    'Beta'                 = 'Beta Channel'
    'InsiderFast'          = 'Beta Channel'                               # legacy
}

function Resolve-ChannelName {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    # GUID embedded in URL or bare
    if ($Value -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
        $guid = $matches[1].ToLower()
        if ($ChannelMap.ContainsKey($guid)) { return $ChannelMap[$guid] }
        return "Unknown Channel (GUID: $guid)"
    }

    # GPO literal value
    if ($PolicyNameMap.ContainsKey($Value)) { return $PolicyNameMap[$Value] }

    return "Unknown Channel (raw: $Value)"
}

function Get-RegPropertySafe {
    param(
        [string]$Path,
        [string]$Name
    )
    if (-not (Test-Path -Path $Path)) { return $null }

    $getParams = @{
        Path        = $Path
        Name        = $Name
        ErrorAction = 'SilentlyContinue'
    }
    $val = Get-ItemProperty @getParams
    if ($null -eq $val) { return $null }
    return $val.$Name
}

# -------------------------------------------------------------------
# Main
# -------------------------------------------------------------------
try {
    # Click-to-Run config (native path, then WOW6432Node fallback)
    $c2rPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration'
    )

    $c2rPath = $null
    foreach ($p in $c2rPaths) {
        if (Test-Path -Path $p) { $c2rPath = $p; break }
    }

    if (-not $c2rPath) {
        Write-Output 'RESULT=NotInstalled'
        Write-Output 'Detail=Microsoft 365 Apps (Click-to-Run) not detected. No Configuration key found in HKLM.'
        Write-Output 'Compliant=N/A'
        exit 0
    }

    $getConfig = @{
        Path        = $c2rPath
        ErrorAction = 'Stop'
    }
    $c2r = Get-ItemProperty @getConfig

    $updateChannel   = $c2r.UpdateChannel
    $cdnBaseUrl      = $c2r.CDNBaseUrl
    $versionToReport = $c2r.VersionToReport
    $platform        = $c2r.Platform
    $productIds      = $c2r.ProductReleaseIds
    $clientCulture   = $c2r.ClientCulture
    $sharedComp      = $c2r.SharedComputerLicensing
    $installRoot     = $c2r.InstallationPath

    $activeChannelName = Resolve-ChannelName -Value $updateChannel
    $cdnChannelName    = Resolve-ChannelName -Value $cdnBaseUrl

    # Policy hive - check GPO then cloud (Intune) policy
    $policyPaths = @(
        'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate',
        'HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate'
    )

    $policyRaw      = $null
    $policyChannel  = $null
    $policySource   = $null
    $policyValueKey = $null

    foreach ($pp in $policyPaths) {
        if (-not (Test-Path -Path $pp)) { continue }

        # Try common value names used by GPO and Intune cloud policy
        foreach ($vn in @('updatebranch','updatepath')) {
            $v = Get-RegPropertySafe -Path $pp -Name $vn
            if ($v) {
                $policyRaw      = $v
                $policyChannel  = Resolve-ChannelName -Value $v
                $policySource   = $pp
                $policyValueKey = $vn
                break
            }
        }
        if ($policyChannel) { break }
    }

    # In-flight migration / drift detection
    $mismatch = $false
    if ($updateChannel -and $cdnBaseUrl -and ($updateChannel -ne $cdnBaseUrl)) {
        $mismatch = $true
    }

    $policyConflict = $false
    if ($policyRaw -and $updateChannel) {
        $pResolved = Resolve-ChannelName -Value $policyRaw
        if ($pResolved -ne $activeChannelName) { $policyConflict = $true }
    }

    # Compliance check against approved channel list
    $compliant = $false
    if ($activeChannelName -and ($AllowedChannels -contains $activeChannelName)) {
        $compliant = $true
    }

    # -------- Output --------
    Write-Output 'RESULT=OK'
    Write-Output "ActiveChannel=$activeChannelName"
    Write-Output "CDNChannel=$cdnChannelName"

    if ($policyChannel) {
        Write-Output "PolicyChannel=$policyChannel"
        Write-Output "PolicySource=$policySource"
        Write-Output "PolicyValueName=$policyValueKey"
    }
    else {
        Write-Output 'PolicyChannel=None (no GPO/Intune channel override detected)'
    }

    Write-Output "Compliant=$compliant"
    Write-Output "AllowedChannels=$($AllowedChannels -join '; ')"
    Write-Output "ChannelMismatch=$mismatch"
    Write-Output "PolicyConflict=$policyConflict"
    Write-Output "Version=$versionToReport"
    Write-Output "Platform=$platform"
    Write-Output "Products=$productIds"
    Write-Output "Culture=$clientCulture"
    if ($sharedComp) { Write-Output "SharedComputerLicensing=$sharedComp" }
    if ($installRoot) { Write-Output "InstallRoot=$installRoot" }
    Write-Output "UpdateChannelRaw=$updateChannel"
    Write-Output "CDNBaseUrlRaw=$cdnBaseUrl"
    if ($policyRaw) { Write-Output "PolicyChannelRaw=$policyRaw" }

    if (-not $compliant) {
        Write-Output 'Note=ActiveChannel is NOT in the approved list for this environment. Flag for remediation.'
    }
    if ($mismatch) {
        Write-Output 'Note=UpdateChannel and CDNBaseUrl differ. A channel change is likely in progress - ODT/GPO flipped UpdateChannel but the Office Automatic Updates 2.0 task has not yet pulled the converted build.'
    }
    if ($policyConflict) {
        Write-Output 'Note=Policy channel differs from ActiveChannel. Either the scheduled task has not yet reconciled, or a cloud policy is overriding a GPO (or vice versa).'
    }

    # -------- Optional: NinjaOne Custom Field --------
    if ($CustomField) {
        $ninjaCmd = Get-Command -Name 'Ninja-Property-Set' -ErrorAction SilentlyContinue
        if ($ninjaCmd) {
            $setParams = @{
                Name  = $CustomField
                Value = $activeChannelName
            }
            Ninja-Property-Set @setParams
            Write-Output "CustomFieldUpdated=$CustomField"
        }
        else {
            Write-Output 'CustomFieldUpdated=SKIPPED (Ninja-Property-Set cmdlet not available in this runspace)'
        }
    }

    # -------- Exit --------
    if ($FailOnNonCompliance -and -not $compliant) {
        exit 2
    }
    exit 0
}
catch {
    Write-Output 'RESULT=Error'
    Write-Output "ErrorMessage=$($_.Exception.Message)"
    Write-Output "ErrorLine=$($_.InvocationInfo.ScriptLineNumber)"
    exit 1
}