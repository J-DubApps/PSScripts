<#
        .SYNOPSIS
            Check if the current user is in a specified group using output from whoami.exe utiil
        .DESCRIPTION
            Does NOT require RSAT ActiveDirectory module or .NET Classes to get true/false group membership
            for current user. Uses search-string method to extract output from whoami.exe, and determine if 
            the specified AD Group was in the whoami.exe output.  
            I had a situation where I needed a non-admin logon script on Windows 10 endpoints to be able to
            check if the current user was a member of a specific AD group, and I couldn't use the 
            WindowsIdentity .NET Class, and is impractical to ever push the RSAT ActiveDirectory PS Module to 
            end-user Windows endpoints.
            This is what I came up with, and it works fairly well but has NOT been tested beyond my own 
            lab and work environments.
            NOTE: I do not believe this works reliably with AD groups with spaces in their names.
            NOTE2 Why MS never put "Kix-like" quick Ingroup checks into Powershell, where it never needs
                to be importing the ActiveDirectory module or using a .NET class -- I will never know.
                I love .NET and all and GPO/GPP is fine - but this one missing feature alone slowed PS 
                logon script adoption  /rant-off :)
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
