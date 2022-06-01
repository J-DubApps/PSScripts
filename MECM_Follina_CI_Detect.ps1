# See this for more info: https://www.cisa.gov/uscert/ncas/current-activity/2022/05/31/microsoft-releases-workaround-guidance-msdt-follina-vulnerability
#
# checks for HKCR\ms-msdt key and reports true/false back to MECM Configuration Item Script
# see MECM_Follina_CI_Detect.ps1 for the remediation action

$msdtcheck = (Get-Item -Path "Registry::HKEY_CLASSES_ROOT\ms-msdt")

if($msdtcheck){
    Write-Output $true
}else{
    Write-Output $false
}
