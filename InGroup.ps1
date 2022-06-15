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
               Simple Check to see if current user is in a specified group using .NET Framework
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
   
   
   #region LICENSE
   <#
         Creative GNU General Public License, version 3 (GPLv3)
         Julian West
         Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
         1. Redistributions of source code must retain the above LICENSE information, this list of conditions and the following disclaimer.
         2. Redistributions in binary form must reproduce the above LICENSE notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
         3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
         THIS SOFTWARE IS PROVIDED BY THE AUTHOR "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE AUTHOR OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
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
         - If you disagree with any of the terms, and any conditions declared: delete it and build your own solution
   #>
   #endregion DISCLAIMER