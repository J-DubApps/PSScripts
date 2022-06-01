# checks for HKCR\ms-msdt key and reports true/false back to MECM Configuration Item Script
# see MECM_Follina_CI_Detect.ps1 for the remediation action

$msdtcheck = (Get-Item -Path "Registry::HKEY_CLASSES_ROOT\ms-msdt")

if($msdtcheck){
    Write-Output $true
}else{
    Write-Output $false
}
