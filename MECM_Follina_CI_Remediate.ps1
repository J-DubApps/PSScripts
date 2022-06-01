# If MECM_Follina_CI_Detect.ps1 returns bool $true during CI/CB pass, this Remediation action nukes the ms-msdt key

Remove-Item -Path "Registry::HKEY_CLASSES_ROOT\ms-msdt" -Recurse -Force -ErrorAction SilentlyContinue
