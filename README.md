# PSScripts

## Description

This repo contains Powershell scripts which are a mix of custom (mine) and forked works of others (attribution provided whenever I have it).

## Scripts

| Script                                                                   | Purpose                                                                                                                                                                              |
| ------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| [Connect-M365Services.ps1](Connect-M365Services.ps1)                     | Modern multi-service connector for Microsoft 365 / Entra ID admin PowerShell (Graph, EXO, SCC, Teams, SharePoint, PnP) with Interactive/DeviceCode/Certificate/ManagedIdentity auth. |
| [Get-AutoStartReport.ps1](Get-AutoStartReport.ps1)                       | Enumerates Windows autostart locations (Registry, Startup folders, logon-triggered Scheduled Tasks, auto-start Services) into a sortable WinForms grid or CSV export.                |
| [Get-UWP-Apps.ps1](Get-UWP-Apps.ps1)                                     | Lists all _available_ Microsoft UWP Appx / MSIX apps on a Win10/11 host (useful for endpoint standardization reviews).                                                               |
| [Get-UWP-Installed.ps1](Get-UWP-Installed.ps1)                           | Lists all _currently installed_ UWP Appx apps on a Win10/11 host (useful for OSD / MDT / SCCM / Intune baselines).                                                                   |
| [InGroup-Whoami.ps1](InGroup-Whoami.ps1)                                 | Checks whether the current user belongs to a named AD group by parsing `whoami.exe` output — no RSAT or .NET AD dependency required.                                                 |
| [InGroup.ps1](InGroup.ps1)                                               | Simple .NET-based function to check if the current user is a member of a specified group.                                                                                            |
| [Ping-IPs.ps1](Ping-IPs.ps1)                                             | Parallel-pings a list of hosts from a CSV file with configurable tries/packet size/wait, optional gridview, stats, and CSV export of results.                                        |
| [Set-Office-Channel-CLI.ps1](Set-Office-Channel-CLI.ps1)                 | Sets the Microsoft 365 Apps update channel via HKLM policy registry (headless/no-UI), with optional control of the in-app channel selector visibility.                               |
| [Set-Office-Channel-GUI.ps1](Set-Office-Channel-GUI.ps1)                 | WinForms GUI equivalent of the CLI channel setter — configure the M365 Apps update channel and constrain/lock the in-app selector.                                                   |
| [Test-AzureCloudShell.ps1](Test-AzureCloudShell.ps1)                     | Returns `$true`/`$false` indicating whether the current PowerShell session is running inside Azure Cloud Shell.                                                                      |
| [get-AutoreplyUsers.ps1](get-AutoreplyUsers.ps1)                         | Exchange Online report of user mailboxes with auto-reply (OOF) enabled, with exclusion support and email-report delivery.                                                            |
| [get-OfficeChannel-njo.ps1](get-OfficeChannel-njo.ps1)                   | NinjaOne-compatible silent reporter for M365 Apps update channel — reports Active/CDN/Policy channels, compliance check, optional custom-field write.                                |
| [MECM/MECM_Follina_CI_Detect.ps1](MECM/MECM_Follina_CI_Detect.ps1)       | MECM Configuration Item detection script for the MSDT "Follina" vulnerability — returns true/false based on presence of the `HKCR\ms-msdt` key.                                      |
| [MECM/MECM_Follina_CI_Remediate.ps1](MECM/MECM_Follina_CI_Remediate.ps1) | MECM remediation paired with the Follina detection CI — removes the `HKCR\ms-msdt` registry key.                                                                                     |

## Disclaimers

- No warranty or guarantee is provided, explicit or implied, for any purpose, use, or adaptation, whether direct or derived, for any code examples or sample data provided on this site.
- USE AT YOUR OWN RISK
- User assumes ANY AND ALL RISK and LIABILITY for any and all usage of examples provided herein. Author assumes no liability for any consequences of using these examples for any purpose whatsoever.
- I make every possible, conceivable, imaginable, effort to indicate the author of each script wherever humanly possible, however, it is possible that I may overlook one or more scripts that I have collected over the years.
- Please let me know if you see something which belongs to you or a solution you have used elsewhere, the URL where it originates, and I will be sure to update my copy to provide clear, absolute, undeniable, irrefutable, inescapable, declaration of the correct author and URL where the original resides.
