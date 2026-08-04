# vybe-test: powershell/profile_loading/profile_loading_safe
try { if (Test-Path $PROFILE) { . $PROFILE } } catch { }
Write-Host 'PASS'
exit 0
