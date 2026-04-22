#Requires -Version 5.1
<#
.SYNOPSIS
    Modern connector for Microsoft 365 / Entra ID administrative PowerShell sessions.

.DESCRIPTION
    Connects to one or more Microsoft 365 administrative services using modern
    authentication - interactive (MFA), certificate-based, managed identity, or
    device code. This is a 2026 rewrite of an older multi-service connector that
    targeted the now-retired MSOnline, AzureAD, and Skype for Business Online
    modules.
    
    MODULE REPLACEMENT MAP (old -> new):
        MSOnline                        ->  Microsoft.Graph
        AzureAD / AzureADPreview        ->  Microsoft.Graph
        SfB Online Connector            ->  MicrosoftTeams (CsOnline* cmdlets)
        SharePointPnPPowerShellOnline   ->  PnP.PowerShell (v2.x / v3.x)
    
    POWERSHELL VERSION GUIDANCE:
        - PowerShell 7.4+ is strongly recommended for all M365 admin work.
        - PnP.PowerShell v3.x REQUIRES PS 7.4.6+. v2.x requires PS 7.2+. There
          is no PS 5.1 path for PnP.
        - Mixing Microsoft.Graph + ExchangeOnlineManagement in the SAME PS 5.1
          session frequently triggers MSAL assembly version conflicts (the
          "Method not found: ... WithBroker(...)" error). PS 7 isolates these
          via Assembly Load Contexts and is the supported fix. Workaround in
          5.1: connect EXO/SCC with -AuthMethod DeviceCode to bypass the WAM
          broker code path, OR run Graph and EXO in separate sessions.
    
    AUTH METHODS:
        Interactive       - Default. Browser-based with MFA.
        DeviceCode        - For headless hosts (SSH, WSL, containers). Also a
                            workaround for the PS 5.1 MSAL/broker conflict.
        Certificate       - Unattended. Requires an app registration with a cert.
        ManagedIdentity   - For Azure-hosted compute (VM, Function, Runbook).
    
    PNP.POWERSHELL APP REGISTRATION:
        PnP.PowerShell v2.x+ retired the shared "PnP Management Shell" multi-
        tenant app registration. You must now register your own app in Entra
        ID before first use. Run the following one time per tenant with a
        Global Admin (and capture the returned ClientId for future use):
            Register-PnPEntraIDAppForInteractiveLogin -ApplicationName 'PnP PowerShell' -Tenant <tenant>.onmicrosoft.com -Interactive

.PARAMETER Services
    One or more services to connect to. Default: All.
    Valid values: Graph, ExchangeOnline, SecurityCompliance, Teams, SharePoint, PnP, All

.PARAMETER AuthMethod
    Interactive (default), DeviceCode, Certificate, ManagedIdentity.

.PARAMETER TenantId
    Target tenant GUID or primary/initial domain. Required for Certificate and
    ManagedIdentity auth. Optional for Interactive/DeviceCode (auto-resolved).

.PARAMETER ClientId
    App registration (service principal) client ID. Required for Certificate auth.

.PARAMETER CertificateThumbprint
    Local cert store thumbprint of the cert registered on the app. Required for
    Certificate auth. Cert must be in CurrentUser\My or LocalMachine\My.

.PARAMETER UserPrincipalName
    Optional hint for interactive auth. Pre-fills the sign-in prompt where the
    underlying module supports it.

.PARAMETER SharePointTenantName
    Tenant name portion of *.sharepoint.com (e.g. 'contoso' for
    contoso.sharepoint.com). If omitted, auto-detected via Graph.

.PARAMETER GraphScopes
    Override the default set of delegated scopes requested for Graph interactive
    auth. Ignored for Certificate / ManagedIdentity (those use app permissions).

.PARAMETER Disconnect
    Disconnects all active M365 sessions and exits.

.PARAMETER SkipModuleCheck
    Skip the prerequisite module validation step. Use when you already know your
    module stack is current.

.PARAMETER ConnectExchangeFirst
    PS 5.1 workaround. Connect EXO/SCC BEFORE Graph to reduce MSAL assembly
    binding conflicts. Sometimes works; PS 7 is the real fix.

.EXAMPLE
    .\Connect-M365Services.ps1
    Interactive MFA connection to all default services.

.EXAMPLE
    .\Connect-M365Services.ps1 -Services Graph,ExchangeOnline
    Interactive connection to Graph and Exchange Online only.

.EXAMPLE
    .\Connect-M365Services.ps1 -Services ExchangeOnline,SecurityCompliance -AuthMethod DeviceCode
    PS 5.1 workaround: device code flow for EXO/SCC bypasses the broker
    extension that triggers MSAL "Method not found" errors.

.EXAMPLE
    $splat = @{
        AuthMethod            = 'Certificate'
        TenantId              = 'contoso.onmicrosoft.com'
        ClientId              = '11111111-2222-3333-4444-555555555555'
        CertificateThumbprint = 'ABCDEF0123456789ABCDEF0123456789ABCDEF01'
        Services              = @('Graph','ExchangeOnline','Teams')
    }
    .\Connect-M365Services.ps1 @splat
    Unattended cert-based connection (typical for scheduled automation).

.EXAMPLE
    .\Connect-M365Services.ps1 -Disconnect
    Tears down all active M365 sessions cleanly.

.NOTES
    Author:       Julian West (original 2022, modernized 2026)
    Version:      2.1
    Target:       PowerShell 5.1+ (PowerShell 7.4+ strongly recommended).
#>

[CmdletBinding()]
param(
    [ValidateSet('Graph','ExchangeOnline','SecurityCompliance','Teams','SharePoint','PnP','All')]
    [string[]]$Services = @('All'),
    
    [ValidateSet('Interactive','DeviceCode','Certificate','ManagedIdentity')]
    [string]$AuthMethod = 'Interactive',
    
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$UserPrincipalName,
    [string]$SharePointTenantName,
    [string[]]$GraphScopes,
    [switch]$Disconnect,
    [switch]$SkipModuleCheck,
    [switch]$ConnectExchangeFirst
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

$script:ModuleRequirements = @{
    'Microsoft.Graph.Authentication'               = [version]'2.15.0'
    'Microsoft.Graph.Identity.DirectoryManagement' = [version]'2.15.0'
    'ExchangeOnlineManagement'                     = [version]'3.4.0'
    'MicrosoftTeams'                               = [version]'5.8.0'
    'Microsoft.Online.SharePoint.PowerShell'       = [version]'16.0.24810.0'
    'PnP.PowerShell'                               = [version]'2.3.0'
}

# PnP.PowerShell minimum PS version (v2.x = 7.2+, v3.x = 7.4.6+)
$script:PnPMinPowerShellVersion = [version]'7.2.0'

$script:DefaultGraphScopes = @(
    'User.Read.All'
    'Group.Read.All'
    'Directory.Read.All'
    'Organization.Read.All'
    'Domain.Read.All'
    'RoleManagement.Read.Directory'
    'AuditLog.Read.All'
    'Policy.Read.All'
)

$script:ServiceModuleMap = @{
    'Graph'              = @('Microsoft.Graph.Authentication','Microsoft.Graph.Identity.DirectoryManagement')
    'ExchangeOnline'     = @('ExchangeOnlineManagement')
    'SecurityCompliance' = @('ExchangeOnlineManagement')
    'Teams'              = @('MicrosoftTeams')
    'SharePoint'         = @('Microsoft.Online.SharePoint.PowerShell')
    'PnP'                = @('PnP.PowerShell')
}

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

function Write-Status {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info','Success','Warning','Error','Step')]
        [string]$Level = 'Info'
    )
    $colors = @{ Info='Cyan'; Success='Green'; Warning='Yellow'; Error='Red'; Step='White' }
    $prefix = @{ Info='[  INFO ]'; Success='[  OK   ]'; Warning='[ WARN  ]'; Error='[ ERROR ]'; Step='[ STEP  ]' }
    Write-Host ("{0} {1}" -f $prefix[$Level], $Message) -ForegroundColor $colors[$Level]
}

function Test-EnvironmentSanity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$TargetServices,
        [Parameter(Mandatory)][string]$AuthMethod
    )
    
    $isPS7  = $PSVersionTable.PSVersion.Major -ge 7
    $psVer  = $PSVersionTable.PSVersion
    $advice = @()
    
    Write-Status -Level Info -Message ("PowerShell version: {0} ({1})" -f $psVer, $PSVersionTable.PSEdition)
    
    if (-not $isPS7) {
        $advice += "PowerShell 5.1 detected. PowerShell 7.4+ is strongly recommended for M365 admin work."
        
        if ($TargetServices -contains 'PnP') {
            Write-Status -Level Error -Message "PnP.PowerShell requires PowerShell 7.2+ (v3.x requires 7.4.6+). It cannot run on PS 5.1."
            Write-Status -Level Warning -Message "PnP will be SKIPPED. Install PS 7 (winget install Microsoft.PowerShell) and re-run from pwsh.exe to use PnP."
            $script:SkipPnP = $true
        }
        
        $hasGraph     = $TargetServices -contains 'Graph'
        $hasExoFamily = ($TargetServices -contains 'ExchangeOnline') -or ($TargetServices -contains 'SecurityCompliance')
        if ($hasGraph -and $hasExoFamily -and $AuthMethod -eq 'Interactive') {
            $advice += "Graph + EXO/SCC in the same PS 5.1 interactive session frequently fails with MSAL 'Method not found: WithBroker' errors."
            $advice += "Workaround: rerun with -AuthMethod DeviceCode for EXO/SCC, or use -ConnectExchangeFirst to flip the load order, or split into two sessions."
        }
    }
    
    # OneDrive module path detection
    $onedrivePaths = $env:PSModulePath -split [System.IO.Path]::PathSeparator |
        Where-Object { $_ -match 'OneDrive' }
    if ($onedrivePaths) {
        Write-Status -Level Warning -Message "PSModulePath includes a OneDrive-synced location:"
        foreach ($p in $onedrivePaths) {
            Write-Host ("    {0}" -f $p) -ForegroundColor DarkYellow
        }
        $advice += "Modules in OneDrive-redirected Documents can fail with Files-On-Demand and slow imports. Consider -Scope AllUsers (requires elevation) to install outside OneDrive."
    }
    
    if ($advice.Count -gt 0) {
        Write-Host ""
        Write-Host "  RECOMMENDATIONS:" -ForegroundColor Yellow
        foreach ($a in $advice) {
            Write-Host ("    - {0}" -f $a) -ForegroundColor Yellow
        }
        Write-Host ""
    }
}

function Test-RequiredModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][version]$MinimumVersion
    )
    
    $installed = Get-Module -ListAvailable -Name $Name |
        Sort-Object Version -Descending |
        Select-Object -First 1
    
    if (-not $installed) {
        Write-Status -Level Warning -Message "$Name not installed."
        return $false
    }
    
    if ($installed.Version -lt $MinimumVersion) {
        Write-Status -Level Warning -Message ("{0} v{1} installed; v{2} or later required." -f $Name, $installed.Version, $MinimumVersion)
        return $false
    }
    
    Write-Verbose ("Module {0} v{1} OK (>= {2})" -f $Name, $installed.Version, $MinimumVersion)
    return $true
}

function Resolve-RequiredModules {
    param(
        [Parameter(Mandatory)][string[]]$TargetServices
    )
    
    $needed = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($svc in $TargetServices) {
        if ($svc -eq 'PnP' -and $script:SkipPnP) { continue }
        foreach ($mod in $script:ServiceModuleMap[$svc]) {
            [void]$needed.Add($mod)
        }
    }
    
    $missing = @()
    foreach ($mod in $needed) {
        $min = $script:ModuleRequirements[$mod]
        if (-not (Test-RequiredModule -Name $mod -MinimumVersion $min)) {
            $missing += [pscustomobject]@{ Name = $mod; MinVersion = $min }
        }
    }
    
    if ($missing.Count -gt 0) {
        Write-Status -Level Warning -Message ("{0} module(s) missing or out of date." -f $missing.Count)
        Write-Host ""
        Write-Host "  Install with:" -ForegroundColor Yellow
        foreach ($m in $missing) {
            Write-Host ("    Install-Module {0} -MinimumVersion {1} -Scope CurrentUser -Force" -f $m.Name, $m.MinVersion) -ForegroundColor Gray
        }
        Write-Host ""
        $resp = Read-Host "Attempt install now? [Y/N]"
        if ($resp -notmatch '^[Yy]') {
            throw "Prerequisite modules missing. Aborting."
        }
        foreach ($m in $missing) {
            Write-Status -Level Step -Message ("Installing {0} v{1}+ ..." -f $m.Name, $m.MinVersion)
            $installParams = @{
                Name           = $m.Name
                MinimumVersion = $m.MinVersion
                Scope          = 'CurrentUser'
                Force          = $true
                AllowClobber   = $true
                ErrorAction    = 'Stop'
            }
            Install-Module @installParams
        }
        Write-Status -Level Success -Message "Module installation complete."
    }
}

function Get-TenantInitialDomain {
    try {
        $domain = Get-MgDomain -ErrorAction Stop | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1
        if ($domain) {
            return ($domain.Id -replace '\.onmicrosoft\.com$','')
        }
    } catch {
        Write-Verbose ("Tenant initial domain lookup failed: {0}" -f $_.Exception.Message)
    }
    return $null
}

function Test-MsalBrokerConflict {
    param([Parameter(Mandatory)][string]$ErrorMessage)
    return ($ErrorMessage -match 'Method not found.*BrokerExtension\.WithBroker' -or
            $ErrorMessage -match 'PublicClientApplicationBuilder.*WithBroker')
}

function Show-MsalConflictGuidance {
    Write-Host ""
    Write-Host "  +-- MSAL ASSEMBLY CONFLICT DETECTED --+" -ForegroundColor Yellow
    Write-Host "  This is the well-known PS 5.1 issue where Microsoft.Graph and EXO" -ForegroundColor Yellow
    Write-Host "  ship incompatible Microsoft.Identity.Client.dll versions and only one" -ForegroundColor Yellow
    Write-Host "  can be loaded into a session at a time. WORKAROUNDS:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    1. Best fix: rerun in PowerShell 7.4+ (pwsh.exe). PS 7 isolates" -ForegroundColor Gray
    Write-Host "       module assemblies via Assembly Load Contexts." -ForegroundColor Gray
    Write-Host ""
    Write-Host "    2. PS 5.1 workaround: rerun EXO/SCC with -AuthMethod DeviceCode," -ForegroundColor Gray
    Write-Host "       which bypasses the broker (WAM) code path entirely:" -ForegroundColor Gray
    Write-Host "         .\Connect-M365Services.ps1 -Services ExchangeOnline,SecurityCompliance -AuthMethod DeviceCode" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    3. PS 5.1 workaround: rerun with -ConnectExchangeFirst to flip load" -ForegroundColor Gray
    Write-Host "       order so EXO's MSAL loads before Graph's." -ForegroundColor Gray
    Write-Host ""
    Write-Host "    4. Run Graph and EXO in separate PowerShell windows." -ForegroundColor Gray
    Write-Host ""
}

# ---------------------------------------------------------------------------
# Per-service connect functions
# ---------------------------------------------------------------------------

function Connect-Graph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AuthMethod,
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [string[]]$Scopes
    )
    
    Write-Status -Level Step -Message "Connecting to Microsoft Graph..."
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    
    $connectArgs = @{ NoWelcome = $true }
    
    switch ($AuthMethod) {
        'Interactive' {
            $connectArgs['Scopes'] = $Scopes
            if ($TenantId) { $connectArgs['TenantId'] = $TenantId }
        }
        'DeviceCode' {
            $connectArgs['Scopes']        = $Scopes
            $connectArgs['UseDeviceCode'] = $true
            if ($TenantId) { $connectArgs['TenantId'] = $TenantId }
        }
        'Certificate' {
            $connectArgs['TenantId']              = $TenantId
            $connectArgs['ClientId']              = $ClientId
            $connectArgs['CertificateThumbprint'] = $CertificateThumbprint
        }
        'ManagedIdentity' {
            $connectArgs['Identity'] = $true
            if ($TenantId) { $connectArgs['TenantId'] = $TenantId }
        }
    }
    
    Connect-MgGraph @connectArgs -ErrorAction Stop
    $ctx = Get-MgContext
    Write-Status -Level Success -Message ("Graph connected. Tenant={0} Account={1} AuthType={2}" -f $ctx.TenantId, $ctx.Account, $ctx.AuthType)
}

function Connect-ExchangeOnlineService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AuthMethod,
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [string]$UserPrincipalName
    )
    
    Write-Status -Level Step -Message "Connecting to Exchange Online (EXO v3 REST)..."
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    
    $connectArgs = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    
    switch ($AuthMethod) {
        'Interactive' {
            if ($UserPrincipalName) { $connectArgs['UserPrincipalName'] = $UserPrincipalName }
        }
        'DeviceCode' {
            $connectArgs['Device'] = $true
        }
        'Certificate' {
            $connectArgs['Organization']          = $TenantId
            $connectArgs['AppId']                 = $ClientId
            $connectArgs['CertificateThumbprint'] = $CertificateThumbprint
        }
        'ManagedIdentity' {
            $connectArgs['ManagedIdentity'] = $true
            $connectArgs['Organization']    = $TenantId
        }
    }
    
    Connect-ExchangeOnline @connectArgs
    Write-Status -Level Success -Message "Exchange Online connected."
}

function Connect-SecurityComplianceService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AuthMethod,
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [string]$UserPrincipalName
    )
    
    Write-Status -Level Step -Message "Connecting to Security & Compliance Center..."
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    
    $connectArgs = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    
    switch ($AuthMethod) {
        'Interactive' {
            if ($UserPrincipalName) { $connectArgs['UserPrincipalName'] = $UserPrincipalName }
        }
        'DeviceCode' {
            $connectArgs['Device'] = $true
        }
        'Certificate' {
            $connectArgs['Organization']          = $TenantId
            $connectArgs['AppId']                 = $ClientId
            $connectArgs['CertificateThumbprint'] = $CertificateThumbprint
        }
        'ManagedIdentity' {
            $connectArgs['ManagedIdentity'] = $true
            $connectArgs['Organization']    = $TenantId
        }
    }
    
    Connect-IPPSSession @connectArgs
    Write-Status -Level Success -Message "Security & Compliance Center connected."
}

function Connect-TeamsService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AuthMethod,
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint
    )
    
    Write-Status -Level Step -Message "Connecting to Microsoft Teams (inc. CsOnline cmdlets)..."
    Import-Module MicrosoftTeams -ErrorAction Stop
    
    $connectArgs = @{ ErrorAction = 'Stop' }
    
    switch ($AuthMethod) {
        'Interactive'     { if ($TenantId) { $connectArgs['TenantId'] = $TenantId } }
        'DeviceCode'      {
            $connectArgs['UseDeviceAuthentication'] = $true
            if ($TenantId) { $connectArgs['TenantId'] = $TenantId }
        }
        'Certificate'     {
            $connectArgs['TenantId']              = $TenantId
            $connectArgs['ApplicationId']         = $ClientId
            $connectArgs['CertificateThumbprint'] = $CertificateThumbprint
        }
        'ManagedIdentity' {
            $connectArgs['Identity'] = $true
            if ($TenantId) { $connectArgs['TenantId'] = $TenantId }
        }
    }
    
    Connect-MicrosoftTeams @connectArgs | Out-Null
    Write-Status -Level Success -Message "Teams connected."
}

function Connect-SharePointOnlineService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AuthMethod,
        [Parameter(Mandatory)][string]$TenantName,
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint
    )
    
    Write-Status -Level Step -Message ("Connecting to SharePoint Online (https://{0}-admin.sharepoint.com)..." -f $TenantName)
    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
    
    $adminUrl = "https://$TenantName-admin.sharepoint.com"
    $connectArgs = @{ Url = $adminUrl; ErrorAction = 'Stop' }
    
    switch ($AuthMethod) {
        'Interactive'     { }
        'DeviceCode'      { Write-Status -Level Warning -Message "Connect-SPOService has no device code flow; falling back to interactive." }
        'Certificate'     {
            if (-not $TenantId) { throw "SharePoint cert auth requires -TenantId" }
            $connectArgs['Tenant']                = $TenantId
            $connectArgs['ApplicationId']         = $ClientId
            $connectArgs['CertificateThumbprint'] = $CertificateThumbprint
        }
        'ManagedIdentity' { throw "Connect-SPOService does not support managed identity. Use PnP instead for that scenario." }
    }
    
    Connect-SPOService @connectArgs
    Write-Status -Level Success -Message "SharePoint Online admin connected."
}

function Connect-PnPOnlineService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AuthMethod,
        [Parameter(Mandatory)][string]$TenantName,
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint
    )
    
    Write-Status -Level Step -Message ("Connecting to PnP Online (https://{0}.sharepoint.com)..." -f $TenantName)
    Import-Module PnP.PowerShell -ErrorAction Stop
    
    $siteUrl = "https://$TenantName.sharepoint.com"
    $connectArgs = @{ Url = $siteUrl; ErrorAction = 'Stop' }
    
    switch ($AuthMethod) {
        'Interactive' {
            $connectArgs['Interactive'] = $true
            if ($ClientId) { $connectArgs['ClientId'] = $ClientId }
        }
        'DeviceCode' {
            $connectArgs['DeviceLogin'] = $true
            if ($ClientId) { $connectArgs['ClientId'] = $ClientId }
        }
        'Certificate' {
            $connectArgs['Tenant']     = $TenantId
            $connectArgs['ClientId']   = $ClientId
            $connectArgs['Thumbprint'] = $CertificateThumbprint
        }
        'ManagedIdentity' {
            $connectArgs['ManagedIdentity'] = $true
        }
    }
    
    try {
        Connect-PnPOnline @connectArgs
        Write-Status -Level Success -Message "PnP Online connected."
    } catch {
        if ($_.Exception.Message -match 'AADSTS700016|application.*not found|was not found in the directory') {
            Write-Status -Level Error -Message "PnP auth failed - likely missing app registration."
            Write-Host ""
            Write-Host "  PnP.PowerShell v2+ requires your own Entra ID app registration." -ForegroundColor Yellow
            Write-Host "  Run ONCE per tenant with Global Admin:" -ForegroundColor Yellow
            Write-Host "    Register-PnPEntraIDAppForInteractiveLogin -ApplicationName 'PnP PowerShell' -Tenant <tenant>.onmicrosoft.com -Interactive" -ForegroundColor Gray
            Write-Host ""
            throw
        }
        throw
    }
}

# ---------------------------------------------------------------------------
# Disconnect
# ---------------------------------------------------------------------------

function Disconnect-AllM365Services {
    Write-Status -Level Step -Message "Disconnecting all M365 sessions..."
    
    try {
        if (Get-Module Microsoft.Graph.Authentication -ListAvailable) {
            if (Get-MgContext -ErrorAction SilentlyContinue) {
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                Write-Status -Level Success -Message "Graph disconnected."
            }
        }
    } catch { Write-Verbose ("Graph disconnect: {0}" -f $_.Exception.Message) }
    
    try {
        if (Get-Module ExchangeOnlineManagement -ListAvailable) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Write-Status -Level Success -Message "Exchange Online / SCC disconnected."
        }
    } catch { Write-Verbose ("EXO disconnect: {0}" -f $_.Exception.Message) }
    
    try {
        if (Get-Module MicrosoftTeams -ListAvailable) {
            Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null
            Write-Status -Level Success -Message "Teams disconnected."
        }
    } catch { Write-Verbose ("Teams disconnect: {0}" -f $_.Exception.Message) }
    
    try {
        if (Get-Module Microsoft.Online.SharePoint.PowerShell -ListAvailable) {
            Disconnect-SPOService -ErrorAction SilentlyContinue
            Write-Status -Level Success -Message "SharePoint Online disconnected."
        }
    } catch { Write-Verbose ("SPO disconnect: {0}" -f $_.Exception.Message) }
    
    try {
        if (Get-Module PnP.PowerShell -ListAvailable) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            Write-Status -Level Success -Message "PnP Online disconnected."
        }
    } catch { Write-Verbose ("PnP disconnect: {0}" -f $_.Exception.Message) }
    
    Get-PSSession | Where-Object {
        $_.ComputerName -match 'outlook\.office365\.com|ps\.compliance\.protection\.outlook\.com'
    } | Remove-PSSession -ErrorAction SilentlyContinue
    
    Write-Status -Level Success -Message "All services disconnected."
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

$script:SkipPnP  = $false
$msalConflictHit = $false

if ($Disconnect) {
    Disconnect-AllM365Services
    return
}

if ($Services -contains 'All') {
    $targetServices = @('Graph','ExchangeOnline','SecurityCompliance','Teams','SharePoint','PnP')
} else {
    $targetServices = $Services
}

switch ($AuthMethod) {
    'Certificate' {
        foreach ($req in 'TenantId','ClientId','CertificateThumbprint') {
            if (-not $PSBoundParameters.ContainsKey($req)) {
                throw "Certificate auth requires -$req"
            }
        }
    }
    'ManagedIdentity' {
        if (-not $PSBoundParameters.ContainsKey('TenantId')) {
            throw "Managed identity auth requires -TenantId"
        }
    }
}

Test-EnvironmentSanity -TargetServices $targetServices -AuthMethod $AuthMethod

if ($script:SkipPnP) {
    $targetServices = $targetServices | Where-Object { $_ -ne 'PnP' }
}

if (-not $GraphScopes) { $GraphScopes = $script:DefaultGraphScopes }

if (-not $SkipModuleCheck) {
    Resolve-RequiredModules -TargetServices $targetServices
}

$needsTenantName = ($targetServices -contains 'SharePoint') -or ($targetServices -contains 'PnP')
if ($needsTenantName -and -not $SharePointTenantName -and ($targetServices -notcontains 'Graph')) {
    Write-Status -Level Warning -Message "SharePoint/PnP requested but no -SharePointTenantName and Graph not in service list."
    $SharePointTenantName = Read-Host "Enter SharePoint tenant name (e.g. 'contoso' for contoso.sharepoint.com)"
}

if ($ConnectExchangeFirst) {
    Write-Status -Level Info -Message "Using -ConnectExchangeFirst load order (EXO before Graph)."
    $connectionOrder = @('ExchangeOnline','SecurityCompliance','Graph','Teams','SharePoint','PnP')
} else {
    $connectionOrder = @('Graph','ExchangeOnline','SecurityCompliance','Teams','SharePoint','PnP')
}

$connected = @()
$failed    = @()

foreach ($svc in $connectionOrder) {
    if ($targetServices -notcontains $svc) { continue }
    
    try {
        switch ($svc) {
            'Graph' {
                $graphParams = @{
                    AuthMethod            = $AuthMethod
                    TenantId              = $TenantId
                    ClientId              = $ClientId
                    CertificateThumbprint = $CertificateThumbprint
                    Scopes                = $GraphScopes
                }
                Connect-Graph @graphParams
                
                if ($needsTenantName -and -not $SharePointTenantName) {
                    $SharePointTenantName = Get-TenantInitialDomain
                    if ($SharePointTenantName) {
                        Write-Status -Level Info -Message ("Resolved SharePoint tenant name: {0}" -f $SharePointTenantName)
                    }
                }
            }
            'ExchangeOnline' {
                $exoParams = @{
                    AuthMethod            = $AuthMethod
                    TenantId              = $TenantId
                    ClientId              = $ClientId
                    CertificateThumbprint = $CertificateThumbprint
                    UserPrincipalName     = $UserPrincipalName
                }
                Connect-ExchangeOnlineService @exoParams
            }
            'SecurityCompliance' {
                $sccParams = @{
                    AuthMethod            = $AuthMethod
                    TenantId              = $TenantId
                    ClientId              = $ClientId
                    CertificateThumbprint = $CertificateThumbprint
                    UserPrincipalName     = $UserPrincipalName
                }
                Connect-SecurityComplianceService @sccParams
            }
            'Teams' {
                $teamsParams = @{
                    AuthMethod            = $AuthMethod
                    TenantId              = $TenantId
                    ClientId              = $ClientId
                    CertificateThumbprint = $CertificateThumbprint
                }
                Connect-TeamsService @teamsParams
            }
            'SharePoint' {
                if (-not $SharePointTenantName) {
                    throw "SharePoint connection needs a tenant name and none could be resolved."
                }
                $spoParams = @{
                    AuthMethod            = $AuthMethod
                    TenantName            = $SharePointTenantName
                    TenantId              = $TenantId
                    ClientId              = $ClientId
                    CertificateThumbprint = $CertificateThumbprint
                }
                Connect-SharePointOnlineService @spoParams
            }
            'PnP' {
                if (-not $SharePointTenantName) {
                    throw "PnP connection needs a tenant name and none could be resolved."
                }
                $pnpParams = @{
                    AuthMethod            = $AuthMethod
                    TenantName            = $SharePointTenantName
                    TenantId              = $TenantId
                    ClientId              = $ClientId
                    CertificateThumbprint = $CertificateThumbprint
                }
                Connect-PnPOnlineService @pnpParams
            }
        }
        $connected += $svc
    } catch {
        $errMsg = $_.Exception.Message
        Write-Status -Level Error -Message ("{0} connection failed: {1}" -f $svc, $errMsg)
        $failed += $svc
        
        if (-not $msalConflictHit -and (Test-MsalBrokerConflict -ErrorMessage $errMsg)) {
            $msalConflictHit = $true
            Show-MsalConflictGuidance
        }
    }
}

Write-Host ""
Write-Host ("=" * 72) -ForegroundColor DarkGray
$connectedDisplay = if ($connected.Count -gt 0) { $connected -join ', ' } else { '(none)' }
Write-Status -Level Info -Message ("Connected services : {0}" -f $connectedDisplay)
if ($failed.Count -gt 0) {
    Write-Status -Level Warning -Message ("Failed services    : {0}" -f ($failed -join ', '))
}
if ($script:SkipPnP) {
    Write-Status -Level Warning -Message "Skipped services   : PnP (requires PowerShell 7.2+)"
}
Write-Host ("=" * 72) -ForegroundColor DarkGray