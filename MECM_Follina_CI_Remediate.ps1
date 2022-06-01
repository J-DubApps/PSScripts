# See this for more info: https://www.cisa.gov/uscert/ncas/current-activity/2022/05/31/microsoft-releases-workaround-guidance-msdt-follina-vulnerability
#
# If MECM_Follina_CI_Detect.ps1 returns bool $true during an MECM CI/CB pass, this Remediation action nukes the ms-msdt key

Remove-Item -Path "Registry::HKEY_CLASSES_ROOT\ms-msdt" -Recurse -Force -ErrorAction SilentlyContinue
