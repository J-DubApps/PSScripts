function InGroup {
    <#
               Name: InGroup.ps1
               05 April 2022
           .NOTES
               Author: Julian West
               Creative GNU General Public License, version 3 (GPLv3);
                ________________________________________________              
               /                                                \             
              |    _________________________________________     |            
              |   |                                         |    |            
              |   |  PS C:\ > WRITE-HOST "$ATTRIBUTION"     |    |            
              |   |                                         |    |            
              |   |         THIS IS A J-DUB SCRIPT          |    |            
              |   |                                         |    |            
              |   |      https://github.com/J-DubAppss      |    |            
              |   |                                         |    |            
              |   |             julianwest.me				|    |   
              |   |       		@julian_west                |    |         
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
           .SYNOPSIS
               Check if the current user is in a specified group
           .DESCRIPTION
               Check if the current user is in a specified group using .NET Framework
               Variations of the PS script exists all over Github in various forms and
               functions, and this is just my own quick take on it.
           .PARAMETER GroupName
               The name of the group to check
           .EXAMPLE
               # Check if the current user is in the Administrators group
               $b = InGroup 'Administrators'
       #>
       Param(
           [string]$GroupName
       )
       
       if($GroupName)
       {
           $mytoken = [System.Security.Principal.WindowsIdentity]::GetCurrent()
           $me = New-Object System.Security.Principal.WindowsPrincipal($mytoken)
           return $me.IsInRole($GroupName)
       }
       else
       {
           $user_token = [System.Security.Principal.WindowsIdentity]::GetCurrent()
           $groups = New-Object System.Collections.ArrayList
           foreach($group in $user_token.Groups)
           {
              [void] $groups.Add( $group.Translate("System.Security.Principal.NTAccount") )
           }
           return $groups
       }
   }