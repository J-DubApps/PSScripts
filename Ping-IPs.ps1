<#
.SYNOPSIS
    Ping simultaneously all the hosts read from a .CSV file.
.DESCRIPTION
    The default action is to ping 5 times each of the hosts simultaneously. 
    The ping results will be shown at the end of the script. The roundtrip time of 
    each successful ping will be used for the calculation of statistics  roundtrip time.
.PARAMETER .Gridview
    Specify this switch to generate gridview at the end of the execution.
.PARAMETER .Flood
    ICMP packet sent once every second if not used. Otherwise, ICMP send continuously 
.PARAMETER .Detail
    To show statistics of the ping tests at the end of the return table.
    The statistics includes:
    N   : number of Tries
    Min : minimum roundtrip time in milliseconds
    AVG : average roundtrip time in milliseconds
    SD  : standard deviation
    Max : maximum roundtrip time in milliseconds
.PARAMETER .Wait
    Time in millisecond for waiting echo reply from destination. Echo reply packet 
    return in longer then the setting will be considered ping failure. 
    Default set to 1000 milliseconds
.PARAMETER .Tries
    Number of ping for each host. Default set to 5. If set tries to 0, the value will set to 4,294,967,295
.PARAMETER .PacketSize
    Buffer size in bytes for ICMP packet. Large package drop indicate protential network equipment issue.
.PARAMETER .Version
    Show the version of the script and quite the script.
.PARAMETER .RelNotice
    Show the release notice and quite the script.
.PARAMETER .Verbose
    Specify this switch to show running time of each phases of the execution of 
    the script.
.PARAMETER .FileName
    Specify file name for hosts in a .CSV file in following format. Default set to .\hosts.csv.
    Columns 'System' and 'Name' are names only. Column 'Host' is used for ping.
    System,Name,Host
    External,www.yahoo.com,www.yahoo.com
    External,www.google.com,www.google.com
    Internal,Dev,192.168.1.10
.PARAMETER .ExportFile
    Export ping results to a file name by appending '_Results_yyyyMMddHHmmss_nnn_' to file name assigned by parameter -FileName.
    
    yyyyMMddHHmmss: date and time of the script started
    nnn:            sequence number start from 000 and increased by one across days
    Ping-IPs.ps1 -FileName hosts.csv -ExportFile
    hosts.csv:                             hosts to be ping
    hosts_Results_yyyyMMddHHmmss_nnn_.csv: created by the script for ping results
.PARAMETER .PingFrequency
    Frequency of ping packet sent. Default set to 1Hz.
.PARAMETER .Reach
    
.EXAMPLE
    Ping-IPs.ps1 -FileName hosts.csv  
    
    2018-12-29 11:46:29.886 AM: All pings started and waiting for their complete.
    2018-12-29 11:46:30.411 AM: @{System=External; Name=www.yahoo.com; Host=www.yahoo.com} ping success.
    2018-12-29 11:46:30.455 AM: @{System=External; Name=www.google.com; Host=www.google.com} ping success.
    2018-12-29 11:46:40.005 AM: All ping tests finished
    System   Name           Host           Loss  Average RTT (ms)
    ------   ----           ----           ----  ----------------
    External www.yahoo.com  www.yahoo.com  0/5  52.7 ms         
    External www.google.com www.google.com 0/5  20 ms           
    Internal Dev            192.168.1.10   10/5 0 ms
.EXAMPLE
    Ping-IPs.ps1 -FileName hosts.csv -Detail
    2018-12-29 11:46:29.886 AM: All pings started and waiting for their complete.
    2018-12-29 11:46:30.411 AM: @{System=External; Name=www.yahoo.com; Host=www.yahoo.com} ping success.
    2018-12-29 11:46:30.455 AM: @{System=External; Name=www.google.com; Host=www.google.com} ping success.
    2018-12-29 11:46:40.005 AM: All ping tests finished
    System   Name           Host           Loss  Average RTT (ms)
    ------   ----           ----           ----  ----------------
    External www.yahoo.com  www.yahoo.com  0/5  52.7 ms         
    External www.google.com www.google.com 0/5  20 ms           
    Internal Dev            192.168.1.10   10/5 0 ms         
.EXAMPLE
    Ping-IPs.ps1 -FileName hosts.csv -ShowTime
    2018-12-29 11:55:28.973 AM: All pings started and waiting for their complete.
    2018-12-29 11:55:29.478 AM: @{System=External; Name=www.yahoo.com; Host=www.yahoo.com} ping success.
    2018-12-29 11:55:29.482 AM: @{System=External; Name=www.google.com; Host=www.google.com} ping success.
    2018-12-29 11:55:39.535 AM: All ping tests finished
    Time                        System   Name           Host           Loss  Average RTT (ms)
    ----                        ------   ----           ----           ----  ----------------
    2018-12-29 11:55:39.018 AM External www.yahoo.com  www.yahoo.com  0/5  52.5 ms         
    2018-12-29 11:55:39.067 AM External www.google.com www.google.com 0/5  20.6 ms         
    2018-12-29 11:55:39.153 AM Internal Dev            192.168.1.10   10/5 0 ms      
.EXAMPLE
    ping-IPs.ps1 -FileName hosts.csv -Gridview
.EXAMPLE
    ping-IPs.ps1 -FileName hosts.csv -Verbose
    2018-12-29 11:59:25.888 AM: Script Ping-IPs.ps1 started
    2018-12-29 11:59:25.898 AM: [Verbose] Create running spool.
    2018-12-29 11:59:25.936 AM: [Verbose] Create host list.
    2018-12-29 11:59:25.968 AM: [Verbose] Initialize ping tests.
    2018-12-29 11:59:26.040 AM: All pings started and waiting for their complete.
    2018-12-29 11:59:26.065 AM: [Verbose] All pings started and waiting for their complete. Waiting.
    2018-12-29 11:59:26.577 AM: @{System=External; Name=www.yahoo.com; Host=www.yahoo.com} ping success.
    2018-12-29 11:59:26.601 AM: @{System=External; Name=www.google.com; Host=www.google.com} ping success.
    .........
    2018-12-29 11:59:31.215 AM: [Verbose] 5.31779s All jobs completed!
    2018-12-29 11:59:31.218 AM: All ping tests finished
    System   Name           Host           Loss Average RTT (ms)
    ------   ----           ----           ---- ----------------
    External www.yahoo.com  www.yahoo.com  0/5  53.2 ms         
    External www.google.com www.google.com 0/5  20.8 ms         
    Internal Dev            192.168.1.10   5/5  0 ms 
.EXAMPLE
    ping-IPs.ps1 -FileName hosts.csv -PacketSize 1500
    2018-12-29 12:10:46.169 PM: All pings started and waiting for their complete.
    2018-12-29 12:10:46.760 PM: @{System=Internal; Name=Dev; Host=192.168.1.10} ping success.
    2018-12-29 12:10:51.784 PM: All ping tests finished
    System   Name           Host           Loss Average RTT (ms)
    ------   ----           ----           ---- ----------------
    External www.yahoo.com  www.yahoo.com  5/5  0 ms            
    External www.google.com www.google.com 5/5  0 ms            
    Internal Dev            192.168.1.10   0/5  3.4 ms       
.EXAMPLE
    PS C:\Users\Damon\desktop> .\Ping-IPs.ps1 -ExportFile 
    2018-12-29 15:17:10.601 PM: All pings started and waiting for their complete.
    2018-12-29 15:17:11.109 PM: @{System=External; Name=www.yahoo.com; Host=www.yahoo.com} ping success.
    2018-12-29 15:17:11.139 PM: @{System=External; Name=www.google.com; Host=www.google.com} ping success.
    2018-12-29 15:17:11.149 PM: @{System=Internal; Name=Dev; Host=192.168.1.10} ping success.
    2018-12-29 15:17:16.195 PM: All ping tests finished
    System   Name           Host           Loss Average RTT (ms)
    ------   ----           ----           ---- ----------------
    External www.yahoo.com  www.yahoo.com  0/5  54 ms           
    External www.google.com www.google.com 0/5  20 ms           
    Internal Dev            192.168.1.10   0/5  0 ms            
    PS C:\Users\Damon\desktop> dir *.csv
        Directory: C:\Users\Damon\desktop
    Mode                LastWriteTime         Length Name                                                                                                                                                                                                            
    ----                -------------         ------ ----                                                                                                                                                                                                            
    -a----       29/12/2018   2:12 AM            121 hosts.csv                                                                                                                                                                                                       
    -a----       29/12/2018   3:17 PM            560 hosts_results_20181229151703_000_.csv                                                                                                                                                                                               
    PS C:\Users\Damon\desktop> type .\hosts_results_20181229151703_000_.csv
    2018-12-29 15:17:11.130 PM,@{System=External; Name=www.yahoo.com; Host=www.yahoo.com},success
    2018-12-29 15:17:11.143 PM,@{System=External; Name=www.google.com; Host=www.google.com},success
    2018-12-29 15:17:11.152 PM,@{System=Internal; Name=Dev; Host=192.168.1.10},success
.NOTES
    Author: Damon
    Date:   December 29, 2018    
#>
 
Param (
   [switch]$Verbose,
   [switch]$Gridview,
   [switch]$Detail,
   [switch]$Version,
   [switch]$RelNotice,
   [switch]$Flood,
   [switch]$ExportFile,
   [ValidateScript({
      if(![bool]([uint16]($_ -as [uint16] -is [uint16]))){
         throw "Invlide value for timeout time of ICMP traffic."
      }
      return $true
   })]
   [string]$Wait="1000",
   [ValidateScript({
      if(![bool]([uint16]($_ -as [uint16] -is [uint16]))){
         throw "Invalide number of tries" 
      } else {
         if( [convert]::toUint16($_,10) -lt 0 ) {
            throw "Value for reach must equal to or greater than 0"
         }
      }
      return $true
   })]
   [string]$Tries="5",
   [ValidateScript({
      if(![bool]([uint16]($_ -as [uint16] -is [uint16]))){
         throw "Invalide packet size assigned" 
      } else {
         if( [convert]::toUint16($_,10) -le 0 ) {
            throw "Value for reach must equal to or greater than 0"
         }
      }
      return $true
   })]
   [uint16]$PacketSize=32,
   [ValidateScript({
      if(-Not ($_ | Test-Path) ){
         throw "File or folder does not exist" 
      }
      if(-Not ($_ | Test-Path -PathType Leaf) ){
         throw "The Path argument must be a file. Folder paths are not allowed."
      }
      return $true
   })]
   [System.IO.FileInfo]$FileName="./hosts.csv",
   [ValidateScript({
      if(![bool]([double]($_ -as [double] -is [double]))) {
         throw "Invalide ping frequency assigned." 
      } else { 
         if([convert]::ToDouble($_) -gt 1) {
            throw "Ping frequency cannot higher than 1."
         }
      }
      return $true
   })]
   [string]$PingFrequency="1",
   [ValidateScript({
      if(![bool]([uint16]($_ -as [uint16] -is [uint16]))){
         throw "Invalide number of reach" 
      } else {
         if( [convert]::toUint16($_,10) -le 0 ) {
            throw "Value for reach must equal to or greater than 1"
         }
      }
      return $true
   })]
   [string]$Reach="3"
)

if(-Not ($FileName | Test-Path) ){
   throw "File or folder does not exist" 
}
if(-Not ($FileName | Test-Path -PathType Leaf) ){
   throw "The Path argument must be a file. Folder paths are not allowed."
}

$DateTimeFormat = "yyyy'-'MM'-'dd HH':'mm':'ss'.'fff tt"
$PingStatusDescription=@{}
$PingStatusDescription[0]="failure"
$PingStatusDescription[1]="success"

if ($Verbose){Write-Host ("{0}: Script {1} started" -f (Get-Date -Format $DateTimeFormat), $MyInvocation.MyCommand.Name)}
if ($Version -or $RelNotice){Write-Host ("{0} Version 1" -f $MyInvocation.MyCommand.Name)}
if ($RelNotice){
   Write-Host ("...")
   Exit
}
if ($Version -or $RelNotice){ Exit }

$StopWatch = [system.diagnostics.stopwatch]::startNew()

$Throttle = 100 #threads

if($Verbose){ Write-Host("{0}: converting parameters" -f (Get-Date -Format $DateTimeFormat)) }
[uint16]$nTries = 5
[uint16]$nWait = 1000
[double]$nPingFrequency = 1
[uint16]$nReach = 3
$nTries = [convert]::ToUInt32($Tries,10)
$nPingCycle = 1000 / [convert]::ToDouble($PingFrequency)
if($nPingCycle -gt 200) {
   $nWait = $nPingCycle - 100 # [convert]::ToUInt32($Wait,10)
} else {
   $nPingCycle = 200
   $nWait = 100
}
$nReachMask = [Math]::Pow(2,[convert]::touint16($Reach)) - 1

if($Verbose){ Write-Host("{0}: frequency = {1} per second" -f (Get-Date -Format $DateTimeFormat), $PingFrequency) }
if($Verbose){ Write-Host("{0}: tries = {1} time(s)" -f (Get-Date -Format $DateTimeFormat), $nTries) }
if($Verbose){ Write-Host("{0}: ping cycle = {1} ms" -f (Get-Date -Format $DateTimeFormat), $nPingCycle) }
if($Verbose){ Write-Host("{0}: packet timeout time = {1} ms" -f (Get-Date -Format $DateTimeFormat), $nWait) }

# script block to ping a host and return an object of Host & Conn
$TestScriptBlock = {
   Param (
      # [string]$TargetHost
      $TargetHost,
      [uint16]$Tries,
      $Wait,
      $Detail,
      $Flood,
      $PingHistory,
      [uint16]$PacketSize = 32,
      $nPingCycle
   )
   $ICMPClient = New-Object System.Net.NetworkInformation.Ping
   $PingOptions = New-Object System.Net.NetworkInformation.PingOptions
   $PingOptions.DontFragment = $True
   $nos = 0
   $not = 0
   $tt  = 0
   $ps = [uint16]0
   
   if($Tries -eq 0) {
      $Tries = [uint16]::MaxValue
   }
   $rrts = 1..$Tries | % {
      $pingStartTime = (get-date)
      $PingResult = ($ICMPClient.Send($TargetHost.Host, $Wait, [System.Byte[]]::CreateInstance([System.Byte],$PacketSize)))
      # ??? $not++
      if($PingResult.Status -eq "Success") { 
         $ps = ($ps -shl 1) -bor 1
         $PingHistory[$TargetHost] = $ps
         $nos++
         # write-host $PingResult.RoundtripTime
         $PingResult.RoundtripTime
      } else {
         $ps = $ps -shl 1
         $PingHistory[$TargetHost] = $ps
         # Write-Host "Loss"
         $null
      }
      $pingEndTime = (get-date)
      if(!$Flood) {sleep -Milliseconds ($nPingCycle - (($pingEndTime - $pingStartTime).TotalMilliseconds))}
   }
   $rrtstats = $rrts | Measure-Object -Average -Maximum -Minimum | select Count, Average, Maximum, Minimum
   $popdev = 0
   foreach ($rrt in $rrts) {
      $popdev +=  [math]::pow(($rrt - $rrtstats.Average), 2)            
   }
   if($rrtstats.Count -gt 1) {
      $sd = [math]::sqrt($popdev / ($rrtstats.Count-1))
   } esle {
      $rrtstats.Maximum = ""
      $rrtstats.Minimum = ""
      $rrtstats.Average = ""
      $rrtstats.Count = ""
      $sd = ""
   }
   $rrtsRS = ""
   if(!$Detail) {
      $rrtsRS = [string]::Format("{0} ms", $rrtstats.Average)
   } else {
      $rrtsRS = [string]::Format("{0} / {1:N1} / {2:N1} / {3:N1} / {4:N1}", $rrtstats.Count, $rrtstats.Minimum, $rrtstats.Average, $sd, $rrtstats.Maximum)
   }
   Return New-Object PSObject -Property @{
      Time = Get-Date -Format "yyyy'-'mm'-'dd HH':'HH':'ss.fff tt"
      System = $TargetHost.System
      Name = $TargetHost.Name
      Host = $TargetHost.Host
      Loss = [string]($Tries-$nos)+'/'+$Tries
      RoundtripTimeStats = $rrtsRS
   }
}
 
if ($Verbose){Write-Host ("{0}: [Verbose] Create running spool." -f (Get-Date -Format $DateTimeFormat))}
$PreviousPingHistory = @{}
$PingHistory = [hashtable]::Synchronized(@{})

# create running spool having $Throttle
$SessionState = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
$runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $SessionState, $Host)
$RunspacePool.Open()
# $RunspacePool.SessionStateProxy.SetVariable('Hash',$hash)
$Jobs = @()
 
if ($Verbose){Write-Host ("{0}: [Verbose] Create host list." -f (Get-Date -Format $DateTimeFormat))}

$Seq = 0
# host list
$c = @()
 
# host/ip to ping
if(test-path $FileName){
   $c = import-csv $FileName
} else {
   exit
}

# Create and start the threads for pings
if ($Verbose){Write-Host ("{0}: [Verbose] Initialize ping tests." -f (Get-Date -Format $DateTimeFormat))}
$c | % {
   $PingHistory[$_] = $null
   $Job = [powershell]::Create().AddScript($TestScriptBlock).AddArgument($_).AddArgument($nTries).AddArgument($nWait).AddArgument($Detail).AddArgument($Flood).AddArgument($PingHistory).AddArgument($PacketSize).AddArgument($nPingCycle)
   $Job.RunspacePool = $RunspacePool
   $Jobs += New-Object PSObject -Property @{
      RunNum = $_
      Pipe = $Job
      Result = $Job.BeginInvoke()
   }
}

Write-Host ("{0}: All pings started and waiting for their complete." -f (Get-Date -Format $DateTimeFormat))

if ($Verbose){Write-Host ("{0}: [Verbose] All pings started and waiting for their complete." -f (Get-Date -Format $DateTimeFormat))}

# wait until all ping jobs completed.
$fns = 0
$fd = $(get-date -format "yyyyMMddHHmmss")
$p = '(.+)(\.csv)'
$r = '$1_results_' + $fd + "_" + $fns.ToString('000') + '_$2'
$efn = $FileName -replace $p, $r
$waitingsign = "|/-\"
$waitingpos = 0
$PreviousDate = (get-date)
if(test-path ($FileName -replace $p, $r)){ remove-item -Force -Confirm:$false ($FileName -replace $p, $r)}
While ( $Jobs.Result.IsCompleted -contains $false) {
   write-host -nonewline $("`r$($waitingsign[$waitingpos])")
   $waitingpos = ++$waitingpos % 4
   Start-Sleep -Milliseconds 1000
   foreach($k in $($PingHistory.Keys)) {
      if(((($PreviousPingHistory[$k] -band $nReachMask) -eq 0) -and (($PingHistory[$k] -band $nReachMask) -eq $nReachMask)) -or ((($PreviousPingHistory[$k] -band $nReachMask) -eq $nReachMask) -and (($PingHistory[$k] -band $nReachMask) -eq 0))) {
         Write-Host ("`r{0}: {1} ping {2} reach = {3}." -f (Get-Date -Format $DateTimeFormat), $k, $PingStatusDescription[$PingHistory[$k] -band [uint16]1], [System.Convert]::ToString($PingHistory[$k] -band [uint16]65535,8).padleft(7,'0')) 
         if($ExportFile){ 
            if( ( (get-date).ToString("yyyyMMdd") -ne $PreviousDate.ToString("yyyyMMdd") ) -or ( (get-item $efn -ErrorAction SilentlyContinue).length/1MB) -gt 20 ){
                $fns++
                $r = '$1_results_' + $fd + "_" + $fns.ToString('000') + '_$2'
                $efn = $FileName -replace $p, $r
                $PreviousDate = (get-date)
            }
            ("{0},{1},{2},{3}" -f (Get-Date -Format $DateTimeFormat), $k, $PingStatusDescription[$PingHistory[$k] -band [uint16]1], [System.Convert]::ToString($PingHistory[$k] -band [uint16]65535,8).padleft(6,'0')) | Out-File -Append -FilePath $efn
         }
         $PreviousPingHistory[$k] = $PingHistory[$k]
      }
   }
}
if ($Verbose){Write-Host -nonewline "`r"}
if ($Verbose){Write-Host ("{0}: [Verbose] {1:N5}s All jobs completed!" -f (Get-Date -Format $DateTimeFormat), $StopWatch.Elapsed.TotalSeconds)}

$StopWatch.Stop()

Write-Host ("`r{0}: All ping tests finished" -f (Get-Date -Format $DateTimeFormat))

# print out the results
$Results = @()
ForEach ($Job in $Jobs) {
   $Results += $Job.Pipe.EndInvoke($Job.Result)
}

if($Detail) {
   $StatsLabel = 'N / Min / AVG / SD / Max'
} else {
   $StatsLabel = 'Average RTT (ms)'
}

if ($Gridview){
   # label='System';expression={$_.system}}
   $Results | Select-Object Time, Sytem, Name, Host, Loss, @{label=$StatsLabel; expression={$_.RoundtripTimeStats}} | Out-GridView -Title "Ping Results"
} else {
   $Results | Format-Table Time, System, Name, Host, Loss, @{label=$StatsLabel; expression={$_.RoundtripTimeStats}} -AutoSize
}
