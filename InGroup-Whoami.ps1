<#
            Name: InGroup_Whoami.ps1
            05 April 2022
        .NOTES
            Author: Julian West
            Creative GNU General Public License, version 3 (GPLv3);
                 _______________________________________________             
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
            Code to check if the current user is a member of a specified AD group using output from whoami.exe
        .DESCRIPTION
            Does NOT require RSAT ActiveDirectory module or .NET Classes to get true/false group membership
            for current user. Uses search-string method to extract output from whoami.exe, and determine if 
            the specified AD Group was in the whoami.exe output.  
            I had a situation where I needed a non-admin logon script on Windows 10 endpoints to be able to
            check current user acct membership in a specific AD group, and I couldn't use the 
            WindowsIdentity .NET Class, and is impractical to push the RSAT ActiveDirectory PS Module to 
            thousands of Windows endpoints.
            This is what I came up with, and it works fairly well but has NOT been tested beyond my own 
            lab and work environment.
            # NOTE: I do not believe this works reliably with AD groups with spaces in their names.
        .PARAMETER GroupName
            The name of the group to check
        .EXAMPLE
            # Check if the current user is in the Administrators group
            $b = InGroup 'Administrators'
    #>

    Param(
          [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$GroupName

    )

$InGroup = $null
$search = $null
$AD_Group_Name = $GroupName
$UserDomainName = $env:userdomain + "\"

$test = whoami /groups /fo list | findstr /C:"$AD_Group_Name"  | Out-String
$search = $test

# write-host $search  # let's see the raw output of whoami 

if($search -eq $null){
	write-host "string not found"
	$InGroup = $false
	write-host $InGroup
	Exit}

$search = $search | select-string -simplematch $AD_Group_Name | Out-String

#write-host $search # check how we're looking before further passes on $search

$search = $search.replace($UserDomainName, '')
#$search = $search.trimstart("") -split '\s+'

#write-host $search  # check to confirm User DNS domain & backslash are gone

$search = $search.replace('Group Name:', '')
$search = $search.trimstart("")

$search = $search -replace "`n|`r"

#write-host "Result of Whoami output before compare:		$search"   # show $search after clean-up before we perform a RegEx pattern match 
# note -- if this is empty then findstr above didn't find anything in gpresult output
#write-host "Original AD Group Name comparing against:	 	$AD_Group_Name" #show original name we're comparing against

#$search = $search.trimstart("") -split '\s+'
#$search = $search -replace "",""

$search = $search | select-string -pattern "^$AD_Group_Name$" | Out-String


If(!$Search -eq $AD_Group_Name){
	write-host "String not found"
	$InGroup = $false
	#write-host "Strings match True/False: $InGroup"
	write-host $InGroup
	Exit
	}else{
	$InGroup = $true	
	#write-host "Strings match True/False: $InGroup"
	write-host $InGroup
	}


# write-host "Run Complete"

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