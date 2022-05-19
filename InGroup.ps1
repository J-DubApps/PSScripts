function InGroup {
    <#
           .SYNOPSIS
               Check if the current user is in a specified group
           .DESCRIPTION
               Check if the current user is in a specified group
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