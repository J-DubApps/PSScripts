#Requires -Version 5.1
#Requires -PSEdition Desktop
using namespace System.Management.Automation
using namespace System.Management.Automation.Language
<#
.SYNOPSIS
    Provides a list of all installed UWP Appx apps on a Win10 or Win11 host.
.DESCRIPTION
    Provides a list of all currently installed UWP Appx apps on a Win10 or Win11 host.  Useful for OSD and standardiation in MDT / SCCM / Intune / etc.
    Must be run as an administrator.
.PARAMETER
    --.
.EXAMPLE
    get-UWP-Apps.ps1
.INPUTS
    None.
.OUTPUTS
    List of apps in PS CLI + optional csv export of the app list.
.NOTES
Author: Julian West
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
    https://github.com/J-DubApps/PSScripts/blob/main/get-UWP-Installed.ps1
.COMPONENT
    --
.FUNCTIONALITY
    --
#>
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


# Initialize an empty array to hold the package information
$installedUWPApps = @()

# Get all installed Appx packages
$allAppxPackages = Get-AppxPackage

# Filter out only the UWP apps that are pre-installed
foreach ($appx in $allAppxPackages) {
  if ($appx.Name -like "*Microsoft.*" -or $appx.Name -like "*windows*") {
    $installedUWPApps += $appx
  }
}

# Display the installed UWP apps
$installedUWPApps | Select-Object Name, PackageFullName, InstallLocation | Format-Table -AutoSize

# Optionally, you can export this information to a CSV file
#$installedUWPApps | Select-Object Name, PackageFullName, InstallLocation | Export-Csv -Path "C:\Path\To\Save\installedUWPApps.csv" -NoTypeInformation

exit 0

#region LICENSE
<#
      BSD 3-Clause License

      Copyright (c) 2023, Julian West
      All rights reserved.

      Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
      1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
      2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
      3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

      THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#>
#endregion LICENSE

#region DISCLAIMER
<#
      DISCLAIMER:
      - Use at your own risk, etc.
      - This is open-source software, if you find an issue try to fix it yourself. There is no support and/or warranty in any kind
      - This is a third-party Software
      - The developer of this Software is NOT sponsored by or affiliated with Microsoft Corp (MSFT) or any of its subsidiaries in any way
      - The Software is not supported by Microsoft Corp (MSFT)
      - By using the Software, you agree to the License, Terms, and any Conditions declared and described above
      - If you disagree with any of the Terms, and any Conditions declared: Just delete it and build your own solution
#>
#endregion DISCLAIMER
