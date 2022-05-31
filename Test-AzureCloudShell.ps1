  <#
               Name: Test-AzureCloudShell.ps1
               31 May 2022
           .NOTES
               Author: Julian West
               Creative GNU General Public License, version 3 (GPLv3);
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
              |   | 	       julianwest.me		    |    |   
              |   |       	@julian_west                |    |            
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
               Check if this PS Session is operating within a remote Azure Cloud Shell 
           .DESCRIPTION
               Function to check if the current PS Session is operating within a remote 
               Azure Cloud Shell.  It will return $true / $false depending on the PS
               environment.  Use this Function alone or with a variable of your choice.
               
               NOTE: Depends on Microsoft to continue setting the "OS" entry within 
               $PSVersionTable to contain the word "azure".  If they change that down the line,
               just review & revise the string-selection pattern in the function or use RegEx.
               
               USAGE: No parameters required, Just call the function!
           .EXAMPLE
               # Assign the function to a variable that gets set to $true or $false if the PS
               # session is running within a remote Azure Cloud Shell
               	
               	$IsACS = Test-AzureCloudShell
       #>

Function Test-AzureCloudShell {

    $ACScheck = $PSVersionTable | select-object OS | select-string -pattern "azure" | Out-String

	If(($PSVersionTable.PSEdition.ToString() -eq 'Core') -and ($ACScheck -match 'azure')){
	# Test $PSVersionTable PSEdition & OS outputs for two tell-tale Azure Cloud environment signs
	# Return $true if both output conditions are present, $false if no match.
	
		return $true
    }else{
    	return $false
    }
}
