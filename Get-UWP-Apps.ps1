#Requires -Version 5.1
#Requires -Version 5.1
#Requires -PSEdition Desktop
<#
.SYNOPSIS
    Scans common Windows autostart locations and displays a sortable, searchable GUI report.

.DESCRIPTION
    Enumerates the most common Registry keys, Startup folders, Scheduled Tasks (logon-triggered),
    and Auto-Start Services, then presents findings in a WinForms DataGridView for quick
    10,000-foot triage by IT staff.

    Does NOT require elevation, but running as Administrator will surface additional machine-level
    scheduled tasks and services that a standard user context cannot enumerate.
.PARAMETER ExcludeServices
    Omit auto-start Windows Services from the report.
 
.PARAMETER ExcludeScheduledTasks
    Omit logon-triggered Scheduled Tasks from the report.
 
.EXAMPLE
    .\Get-AutoStartReport.ps1
    Full scan including services and scheduled tasks.
 
.EXAMPLE
    .\Get-AutoStartReport.ps1 -ExcludeServices -ExcludeScheduledTasks
    Registry and Startup folder entries only — quick lightweight view.
 
.EXAMPLE
    .\Get-AutoStartReport.ps1 -ExcludeServices
    Everything except services.
    
.EXAMPLE
    .\Get-AutoStartReport.ps1 -ExportOnly
    Full scan, no GUI — writes CSV to %USERPROFILE%\Downloads.
 
.EXAMPLE
    .\Get-AutoStartReport.ps1 -ExportOnly -ExcludeServices -ExcludeScheduledTasks
    Lightweight headless export for RMM deployment.

.INPUTS
    None.

.OUTPUTS
    List of Startup Items discovered + optional csv export of resulting list.

.NOTES
    Author  : Julian West (with a quick-assist from Claude Code / Anthropic)
    Version : 1.2.0
    Date    : 2026-03-20
    Target  : PowerShell 5.1+ on Windows 10/11
    
               BSD 3-Clause License;
               - see License Region at-end of script for more information
                ________________________________________________
               /                                                \
              |    _________________________________________     |
              |   |                                         |    |
              |   |  PS C:\ > WRITE-HOST $ATTRIBUTION	    |    |
              |   |                                         |    |
              |   |         THIS IS A J-DUB SCRIPT          |    |
              |   |                                         |    |
              |   |      https://github.com/J-DubAppss      |    |
              |   |                                         |    |
              |   | 	       julianwest.me                |    |
              |   |             @julian_west                |    |
              |   |                                         |    |
              |   |                                         |    |
              |   |                                         |    |
              |   |                                         |    |
              |   |                                         |    |
              |   |_________________________________________|    |
              |                                                  |
               \_________________________________________________/
                      \___________________________________/
                   ___________________________________________
                _-'    .-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.  --- `-_
             _-'.-.-. .---.-.-.-.-.-.-.-.-.-.-.-.-.-.-.--.  .-.-.`-_
          _-'.-.-.-. .---.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-`__`. .-.-.-.`-_
       _-'.-.-.-.-. .-----.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-----. .-.-.-.-.`-_
    _-'.-.-.-.-.-. .---.-. .-------------------------. .-.---. .---.-.-.-.`-_
   :-------------------------------------------------------------------------:
   `---._.-------------------------------------------------------------._.---'
.LINK
    https://julianwest.me
.LINK
    https://github.com/J-DubApps/PSScripts/blob/main/Get-AutoStartReport.ps1
.COMPONENT
    --
.FUNCTIONALITY
    --
#>

[CmdletBinding()]
param(
    [switch]$ExcludeServices,
    [switch]$ExcludeScheduledTasks,
    [switch]$ExportOnly
)

if ($ENV:PROCESSOR_ARCHITEW6432 -eq 'AMD64') {
  try {
  &"$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCOMMANDPATH
  }
  catch {
    Throw ('Failed to start {0}' -f $PSCOMMANDPATH)
  }

  exit
}
#endregion ARM64Handling

#region ── Assembly Loading ──────────────────────────────────────────────────────
if (-not $ExportOnly) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
}
#endregion
 
#region ── Helper Functions ──────────────────────────────────────────────────────
 
function Get-RegistryAutostarts {
    <#
    .SYNOPSIS
        Reads values from a single registry path and returns structured objects.
    #>
    [CmdletBinding()]
    param(
        [string]$Path,
        [string]$Category,
        [string]$Scope
    )
 
    if (-not (Test-Path -LiteralPath "Registry::$Path")) { return }
 
    $key = Get-Item -LiteralPath "Registry::$Path" -ErrorAction SilentlyContinue
    if ($null -eq $key) { return }
 
    foreach ($valueName in $key.GetValueNames()) {
        # Skip the default (unnamed) value unless it actually contains something useful
        if ([string]::IsNullOrWhiteSpace($valueName)) { continue }
 
        $valueData = $key.GetValue($valueName, $null, 'DoNotExpandEnvironmentNames')
        if ([string]::IsNullOrWhiteSpace($valueData)) { continue }
 
        [PSCustomObject]@{
            Category = $Category
            Scope    = $Scope
            Name     = $valueName
            Command  = $valueData.ToString().Trim()
            Location = $Path
        }
    }
}
 
function Get-WinlogonValues {
    <#
    .SYNOPSIS
        Reads specific named values from the Winlogon key that can inject autostart commands.
    #>
    [CmdletBinding()]
    param()
 
    $winlogonPath = 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
 
    if (-not (Test-Path -LiteralPath "Registry::$winlogonPath")) { return }
 
    $key = Get-Item -LiteralPath "Registry::$winlogonPath" -ErrorAction SilentlyContinue
    if ($null -eq $key) { return }
 
    $interestingValues = @('Userinit', 'Shell', 'Taskman', 'AppSetup')
 
    foreach ($valName in $interestingValues) {
        $valData = $key.GetValue($valName, $null)
        if ([string]::IsNullOrWhiteSpace($valData)) { continue }
 
        # Flag if the value deviates from known defaults
        $isDefault = $false
        switch ($valName) {
            'Userinit' { $isDefault = ($valData.Trim().TrimEnd(',') -eq 'C:\Windows\system32\userinit.exe') }
            'Shell'    { $isDefault = ($valData.Trim() -eq 'explorer.exe') }
        }
 
        $displayName = if ($isDefault) { "$valName [DEFAULT]" } else { "$valName [MODIFIED]" }
 
        [PSCustomObject]@{
            Category = 'Winlogon'
            Scope    = 'Machine'
            Name     = $displayName
            Command  = $valData.ToString().Trim()
            Location = $winlogonPath
        }
    }
}
 
function Get-StartupFolderItems {
    <#
    .SYNOPSIS
        Enumerates shortcuts and executables in the user and common Startup folders.
    #>
    [CmdletBinding()]
    param()
 
    $folders = @(
        @{
            Path  = [Environment]::GetFolderPath('Startup')
            Scope = 'User'
        }
        @{
            Path  = [Environment]::GetFolderPath('CommonStartup')
            Scope = 'Machine'
        }
    )
 
    $shell = $null
    try {
        $shell = New-Object -ComObject WScript.Shell
    }
    catch {
        Write-Warning "Could not create WScript.Shell COM object for shortcut resolution."
    }
 
    foreach ($folder in $folders) {
        if (-not (Test-Path -LiteralPath $folder.Path)) { continue }
 
        $items = Get-ChildItem -LiteralPath $folder.Path -File -ErrorAction SilentlyContinue
 
        foreach ($item in $items) {
            # Skip desktop.ini
            if ($item.Name -eq 'desktop.ini') { continue }
 
            $command = $item.FullName
 
            # Resolve .lnk shortcut targets
            if ($item.Extension -eq '.lnk' -and $null -ne $shell) {
                try {
                    $shortcut = $shell.CreateShortcut($item.FullName)
                    $target = $shortcut.TargetPath
                    $args   = $shortcut.Arguments
                    if (-not [string]::IsNullOrWhiteSpace($target)) {
                        $command = if ([string]::IsNullOrWhiteSpace($args)) { $target } else { "$target $args" }
                    }
                }
                catch {
                    # Fall back to showing the .lnk path itself
                }
            }
 
            [PSCustomObject]@{
                Category = 'Startup Folder'
                Scope    = $folder.Scope
                Name     = $item.Name
                Command  = $command
                Location = $folder.Path
            }
        }
    }
}
 
function Get-LogonScheduledTasks {
    <#
    .SYNOPSIS
        Returns scheduled tasks that have a LogonTrigger defined.
    #>
    [CmdletBinding()]
    param()
 
    try {
        $allTasks = Get-ScheduledTask -ErrorAction Stop |
            Where-Object { $_.State -ne 'Disabled' }
    }
    catch {
        Write-Warning "Could not enumerate scheduled tasks: $($_.Exception.Message)"
        return
    }
 
    foreach ($task in $allTasks) {
        $hasLogonTrigger = $false
        foreach ($trigger in $task.Triggers) {
            if ($trigger -is [Microsoft.Management.Infrastructure.CimInstance]) {
                $cimClass = $trigger.CimClass.CimClassName
                if ($cimClass -eq 'MSFT_TaskLogonTrigger') {
                    $hasLogonTrigger = $true
                    break
                }
            }
        }
        if (-not $hasLogonTrigger) { continue }
 
        # Extract the action (Execute + Arguments)
        $actionStr = ''
        foreach ($action in $task.Actions) {
            if ($action.Execute) {
                $exe  = $action.Execute
                $args = $action.Arguments
                $actionStr = if ([string]::IsNullOrWhiteSpace($args)) { $exe } else { "$exe $args" }
                break   # Show the first action only for readability
            }
        }
 
        $scope = if ($task.Principal.UserId -match 'S-1-5-18|SYSTEM|LOCAL SERVICE|NETWORK SERVICE') {
            'Machine'
        }
        else {
            'User'
        }
 
        [PSCustomObject]@{
            Category = 'Scheduled Task (Logon)'
            Scope    = $scope
            Name     = "$($task.TaskPath)$($task.TaskName)"
            Command  = $actionStr
            Location = 'Task Scheduler'
        }
    }
}
 
function Get-AutoStartServices {
    <#
    .SYNOPSIS
        Returns services set to Automatic or Automatic (Delayed Start).
    #>
    [CmdletBinding()]
    param()
 
    $svcParams = @{
        ErrorAction = 'SilentlyContinue'
    }
 
    $services = Get-CimInstance -ClassName Win32_Service @svcParams |
        Where-Object { $_.StartMode -eq 'Auto' }
 
    foreach ($svc in $services) {
        [PSCustomObject]@{
            Category = "Service ($($svc.StartMode))"
            Scope    = 'Machine'
            Name     = "$($svc.DisplayName) [$($svc.Name)]"
            Command  = $svc.PathName
            Location = "HKLM\SYSTEM\CurrentControlSet\Services\$($svc.Name)"
        }
    }
}
 
#endregion
 
#region ── Data Collection ───────────────────────────────────────────────────────
 
Write-Host 'Scanning autostart locations...' -ForegroundColor Cyan
 
$results = [System.Collections.Generic.List[PSCustomObject]]::new()
 
# ── Registry Run Keys ──
 
$registryTargets = @(
    # User-scope Run keys
    @{ Path = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Run';                              Category = 'Run';        Scope = 'User' }
    @{ Path = 'HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce';                          Category = 'RunOnce';    Scope = 'User' }
    @{ Path = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run';            Category = 'Policy Run'; Scope = 'User' }
 
    # Machine-scope Run keys
    @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run';                              Category = 'Run';        Scope = 'Machine' }
    @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce';                          Category = 'RunOnce';    Scope = 'Machine' }
    @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run';            Category = 'Policy Run'; Scope = 'Machine' }
 
    # WOW6432Node (32-bit on 64-bit)
    @{ Path = 'HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run';                  Category = 'Run (x86)';  Scope = 'Machine' }
    @{ Path = 'HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce';              Category = 'RunOnce (x86)'; Scope = 'Machine' }
 
    # RunServices (legacy but still checked)
    @{ Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices';                      Category = 'RunServices'; Scope = 'Machine' }
    @{ Path = 'HKCU\Software\Microsoft\Windows\CurrentVersion\RunServices';                      Category = 'RunServices'; Scope = 'User' }
 
    # Shell-related per-user
    @{ Path = 'HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows';                       Category = 'NT Windows'; Scope = 'User' }
)
 
foreach ($target in $registryTargets) {
    $splat = @{
        Path     = $target.Path
        Category = $target.Category
        Scope    = $target.Scope
    }
    $items = Get-RegistryAutostarts @splat
    if ($items) { $results.AddRange([PSCustomObject[]]@($items)) }
}
 
# ── Winlogon special values ──
$winlogonItems = Get-WinlogonValues
if ($winlogonItems) { $results.AddRange([PSCustomObject[]]@($winlogonItems)) }
 
# ── Startup Folders ──
$startupItems = Get-StartupFolderItems
if ($startupItems) { $results.AddRange([PSCustomObject[]]@($startupItems)) }
 
# ── Scheduled Tasks with Logon Triggers ──
if (-not $ExcludeScheduledTasks) {
    $taskItems = Get-LogonScheduledTasks
    if ($taskItems) { $results.AddRange([PSCustomObject[]]@($taskItems)) }
}
else {
    Write-Host '  Skipping Scheduled Tasks (excluded by parameter).' -ForegroundColor DarkGray
}
 
# ── Auto-Start Services ──
if (-not $ExcludeServices) {
    $serviceItems = Get-AutoStartServices
    if ($serviceItems) { $results.AddRange([PSCustomObject[]]@($serviceItems)) }
}
else {
    Write-Host '  Skipping Services (excluded by parameter).' -ForegroundColor DarkGray
}
 
# ── Summary ──
$excludedParts = @()
if ($ExcludeServices)       { $excludedParts += 'Services' }
if ($ExcludeScheduledTasks) { $excludedParts += 'Scheduled Tasks' }
$excludeNote = if ($excludedParts.Count -gt 0) { "  (Excluded: $($excludedParts -join ', '))" } else { '' }
Write-Host "Found $($results.Count) autostart entries.$excludeNote" -ForegroundColor Green
 
#endregion
 
#region ── Output: Headless CSV or GUI ───────────────────────────────────────────
 
if ($ExportOnly) {
    # ── Headless CSV Export ──
    $computerName = $env:COMPUTERNAME
    $downloadsPath = Join-Path -Path $env:USERPROFILE -ChildPath 'Downloads'
 
    # Fallback if Downloads doesn't exist (e.g., server core, service account)
    if (-not (Test-Path -LiteralPath $downloadsPath)) {
        $downloadsPath = $env:TEMP
        Write-Warning "Downloads folder not found. Falling back to: $downloadsPath"
    }
 
    $csvFileName = "AutostartReport_${computerName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $csvPath     = Join-Path -Path $downloadsPath -ChildPath $csvFileName
 
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported $($results.Count) entries to: $csvPath" -ForegroundColor Green
 
    # Emit the path as output for RMM tools to capture
    Write-Output $csvPath
    return
}
 
# ── Interactive GUI (default) ──
 
$computerName = $env:COMPUTERNAME
$userName     = $env:USERNAME
$osCaption    = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
$scanTime     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$isAdmin      = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
                    [Security.Principal.WindowsBuiltInRole]::Administrator)
$elevationTag = if ($isAdmin) { 'Elevated' } else { 'Standard (run as Admin for full results)' }
$excludeBanner = if ($excludedParts.Count -gt 0) { "   |   Excluded: $($excludedParts -join ', ')" } else { '' }
 
# ── Main Form ──
 
$form = New-Object System.Windows.Forms.Form
$formProps = @{
    Text            = "Autostart Report  |  $computerName  |  $scanTime"
    Size            = New-Object System.Drawing.Size(1280, 720)
    StartPosition   = 'CenterScreen'
    MinimumSize     = New-Object System.Drawing.Size(900, 500)
    Font            = New-Object System.Drawing.Font('Segoe UI', 9)
}
foreach ($prop in $formProps.GetEnumerator()) { $form.$($prop.Key) = $prop.Value }
 
# ── Info Banner ──
 
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfoProps = @{
    Text      = "  Computer: $computerName   |   User: $userName   |   OS: $osCaption   |   Context: $elevationTag   |   Entries: $($results.Count)$excludeBanner"
    Dock      = 'Top'
    Height    = 32
    BackColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
    ForeColor = [System.Drawing.Color]::White
    TextAlign = 'MiddleLeft'
    Font      = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
}
foreach ($prop in $lblInfoProps.GetEnumerator()) { $lblInfo.$($prop.Key) = $prop.Value }
$form.Controls.Add($lblInfo)
 
# ── Filter Bar Panel ──
 
$pnlFilter = New-Object System.Windows.Forms.Panel
$pnlFilter.Dock   = 'Top'
$pnlFilter.Height = 40
$pnlFilter.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
$form.Controls.Add($pnlFilter)
 
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text     = 'Filter:'
$lblSearch.AutoSize = $true
$lblSearch.Location = New-Object System.Drawing.Point(10, 10)
$pnlFilter.Controls.Add($lblSearch)
 
$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(55, 7)
$txtFilter.Size     = New-Object System.Drawing.Size(300, 24)
$pnlFilter.Controls.Add($txtFilter)
 
$cmbCategory = New-Object System.Windows.Forms.ComboBox
$cmbCategory.Location      = New-Object System.Drawing.Point(370, 7)
$cmbCategory.Size           = New-Object System.Drawing.Size(180, 24)
$cmbCategory.DropDownStyle  = 'DropDownList'
$cmbCategory.Items.Add('All Categories') | Out-Null
$categories = $results | Select-Object -ExpandProperty Category -Unique | Sort-Object
foreach ($cat in $categories) { $cmbCategory.Items.Add($cat) | Out-Null }
$cmbCategory.SelectedIndex = 0
$pnlFilter.Controls.Add($cmbCategory)
 
$cmbScope = New-Object System.Windows.Forms.ComboBox
$cmbScope.Location      = New-Object System.Drawing.Point(560, 7)
$cmbScope.Size           = New-Object System.Drawing.Size(120, 24)
$cmbScope.DropDownStyle  = 'DropDownList'
$cmbScope.Items.AddRange(@('All Scopes', 'User', 'Machine'))
$cmbScope.SelectedIndex = 0
$pnlFilter.Controls.Add($cmbScope)
 
$btnExport = New-Object System.Windows.Forms.Button
$btnExportProps = @{
    Text      = 'Export CSV'
    Location  = New-Object System.Drawing.Point(700, 6)
    Size      = New-Object System.Drawing.Size(90, 26)
    FlatStyle = 'Flat'
}
foreach ($prop in $btnExportProps.GetEnumerator()) { $btnExport.$($prop.Key) = $prop.Value }
$pnlFilter.Controls.Add($btnExport)
 
# ── DataGridView ──
 
$dgv = New-Object System.Windows.Forms.DataGridView
$dgvProps = @{
    Dock                       = 'Fill'
    ReadOnly                   = $true
    AllowUserToAddRows         = $false
    AllowUserToDeleteRows      = $false
    AllowUserToResizeRows      = $false
    AutoSizeColumnsMode        = 'Fill'
    SelectionMode              = 'FullRowSelect'
    RowHeadersVisible          = $false
    BackgroundColor            = [System.Drawing.Color]::White
    BorderStyle                = 'None'
    AlternatingRowsDefaultCellStyle = (New-Object System.Windows.Forms.DataGridViewCellStyle -Property @{
        BackColor = [System.Drawing.Color]::FromArgb(245, 248, 250)
    })
    ColumnHeadersDefaultCellStyle = (New-Object System.Windows.Forms.DataGridViewCellStyle -Property @{
        BackColor = [System.Drawing.Color]::FromArgb(52, 58, 64)
        ForeColor = [System.Drawing.Color]::White
        Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    })
    EnableHeadersVisualStyles  = $false
}
foreach ($prop in $dgvProps.GetEnumerator()) { $dgv.$($prop.Key) = $prop.Value }
 
# Define columns
$colDefs = @(
    @{ Name = 'Category'; Header = 'Category';  Width = 140 }
    @{ Name = 'Scope';    Header = 'Scope';     Width = 70  }
    @{ Name = 'Name';     Header = 'Name';      Width = 250 }
    @{ Name = 'Command';  Header = 'Command';   Width = 400 }
    @{ Name = 'Location'; Header = 'Location';  Width = 350 }
)
 
foreach ($colDef in $colDefs) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name             = $colDef.Name
    $col.HeaderText       = $colDef.Header
    $col.FillWeight       = $colDef.Width
    $col.SortMode         = 'Automatic'
    $dgv.Columns.Add($col) | Out-Null
}
 
$form.Controls.Add($dgv)
 
# ── Status Bar ──
 
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready  |  $($results.Count) entries loaded"
$statusBar.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusBar)
 
#endregion
 
#region ── Data Binding & Filtering ──────────────────────────────────────────────
 
function Update-GridData {
    <#
    .SYNOPSIS
        Applies the current filter/category/scope selections and refreshes the grid.
    #>
    $dgv.Rows.Clear()
 
    $filterText = $txtFilter.Text.Trim()
    $catFilter  = $cmbCategory.SelectedItem.ToString()
    $scopeFilter = $cmbScope.SelectedItem.ToString()
 
    $filtered = $results
 
    # Category filter
    if ($catFilter -ne 'All Categories') {
        $filtered = $filtered | Where-Object { $_.Category -eq $catFilter }
    }
 
    # Scope filter
    if ($scopeFilter -ne 'All Scopes') {
        $filtered = $filtered | Where-Object { $_.Scope -eq $scopeFilter }
    }
 
    # Free-text filter (search across all fields)
    if (-not [string]::IsNullOrWhiteSpace($filterText)) {
        $filtered = $filtered | Where-Object {
            ($_.Category -like "*$filterText*") -or
            ($_.Name     -like "*$filterText*") -or
            ($_.Command  -like "*$filterText*") -or
            ($_.Location -like "*$filterText*")
        }
    }
 
    foreach ($item in $filtered) {
        $dgv.Rows.Add(
            $item.Category,
            $item.Scope,
            $item.Name,
            $item.Command,
            $item.Location
        ) | Out-Null
    }
 
    $statusLabel.Text = "Showing $($dgv.Rows.Count) of $($results.Count) entries"
}
 
# Wire up filter events
$txtFilter.Add_TextChanged({ Update-GridData })
$cmbCategory.Add_SelectedIndexChanged({ Update-GridData })
$cmbScope.Add_SelectedIndexChanged({ Update-GridData })
 
# Export button
$btnExport.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter   = 'CSV Files (*.csv)|*.csv'
    $saveDialog.FileName = "AutostartReport_${computerName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
 
    if ($saveDialog.ShowDialog() -eq 'OK') {
        $results | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show(
            "Exported $($results.Count) entries to:`n$($saveDialog.FileName)",
            'Export Complete',
            'OK',
            'Information'
        )
    }
})
 
# Right-click context menu for copying cell data
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$copyItem = $contextMenu.Items.Add('Copy Cell Value')
$copyItem.Add_Click({
    if ($dgv.CurrentCell -and $dgv.CurrentCell.Value) {
        [System.Windows.Forms.Clipboard]::SetText($dgv.CurrentCell.Value.ToString())
    }
})
$copyRow = $contextMenu.Items.Add('Copy Full Row')
$copyRow.Add_Click({
    if ($dgv.CurrentRow) {
        $rowText = ($dgv.CurrentRow.Cells | ForEach-Object { $_.Value }) -join "`t"
        [System.Windows.Forms.Clipboard]::SetText($rowText)
    }
})
$dgv.ContextMenuStrip = $contextMenu
 
# Initial data load
Update-GridData
 
#endregion
 
#region ── Show Form ─────────────────────────────────────────────────────────────
 
# Bring the grid into focus so column sorting works immediately
$form.Add_Shown({ $dgv.Focus() })
 
[void]$form.ShowDialog()
 
# Clean up COM object
if ($null -ne $shell) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
}
 
#endregion
