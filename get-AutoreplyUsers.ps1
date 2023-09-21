#Requires -Version 5.1
#Requires -PSEdition Desktop
using namespace System.Management.Automation
using namespace System.Management.Automation.Language
<#
.SYNOPSIS
    Provides current EXO user mailboxes with AutoReply out-of-office enabled.
.DESCRIPTION
    Provides a list and emailed report of EXO user mailboxes with AutoReply enabled.  Allows for excluding certain mailboxes based on UPN or Identity
.PARAMETER
    --.
.EXAMPLE
    get-AutoReplyUsers.ps1
.INPUTS
    None.
.OUTPUTS
    $autoReplyInfo result + email report.
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
    https://github.com/J-DubApps/PS-MANAGE/CHANGELOG.TXT
.LINK
    https://github.com/J-DubApps/PS-MANAGE
.COMPONENT
    --
.FUNCTIONALITY
    --
#>
if ($ENV:PROCESSOR_ARCHITEW6432 -eq 'AMD64')
{
   try
   {
      &"$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCOMMANDPATH
   }
   catch
   {
      Throw ('Failed to start {0}' -f $PSCOMMANDPATH)
   }

   exit
}
#endregion ARM64Handling


# Import required modules
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName <YourAdminAccount>

# Initialize an empty array to store mailbox info
$autoReplyInfo = @()

# NOTE: Some mailboxes may simply produce errors as being not-locatable in EXO (due to perhaps being licensed, but
# not having a mailbox on-prem or in EXO).  There may also be certain mailboxes with a GUID Identity, without a name, which we also wish to exclude.
# The next two array variables below will exclude searching / outputting specific UPNs or Identities. 

# Define an array of UPNs and Identities to exclude
$excludeUPNs = @("user_upn1@domain.com", "user_upn2@domain.com", "user_upn3@domain.com")
$excludeIdentities = @("Identity1", "GUID1", "GUID2") # Some identities may just be a GUID, instead of the usual Display name

# Get all enabled and licensed Office 365 users
$users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true -and $_.BlockCredential -eq $false }

# Loop through each user to get their auto-reply configuration
foreach ($user in $users) {
    if ($excludeUPNs -notcontains $user.UserPrincipalName) {
        $mailbox = Get-Mailbox -Identity $user.UserPrincipalName -RecipientTypeDetails UserMailbox
        if ($null -ne $mailbox) {
            $autoReplyConfig = $mailbox | Get-MailboxAutoReplyConfiguration

            if ($autoReplyConfig.AutoReplyState -ne "Disabled" -and $excludeIdentities -notcontains $autoReplyConfig.Identity) {
                $info = [PSCustomObject]@{
                    Identity       = $autoReplyConfig.Identity
                    StartTime      = $autoReplyConfig.StartTime
                    EndTime        = $autoReplyConfig.EndTime
                    AutoReplyState = $autoReplyConfig.AutoReplyState
                }

                $autoReplyInfo += $info
            }
        }
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

# Output the auto-reply information
$autoReplyInfo | Select-Object Identity, StartTime, EndTime, AutoReplyState | Format-Table -AutoSize

# Convert the auto-reply information to an HTML table
$htmlTable = $autoReplyInfo | Select-Object Identity, StartTime, EndTime, AutoReplyState | ConvertTo-Html -As Table

# Email settings
$emailFrom = "ps-script@domain.com"
$emailTo = "sysadminjwest@domain.com"
$emailSubject = "Daily Auto-Reply Report"
$smtpServer = "0.0.0.0"  # place your local SMTP on-prem relay IP or FQDN here - NOTE: you may need to whitelist the sending host running this script

# Send the email
Send-MailMessage -From $emailFrom -To $emailTo -Subject $emailSubject -Body $htmlTable -BodyAsHtml -SmtpServer $smtpServer

exit 0

#region LICENSE
<#
      BSD 3-Clause License

      Copyright (c) 2022, Julian West
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
