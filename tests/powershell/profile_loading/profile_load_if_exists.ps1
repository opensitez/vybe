# vybe-test: powershell/profile_loading/profile_load_if_exists
if (Test-Path $PROFILE) {
    . $PROFILE
}
Write-Host 'PASS'
exit 0
