# vybe-test: powershell/profile_loading/try_load_profile
if (Test-Path $PROFILE) {
    . $PROFILE
}
Write-Host 'PASS'
exit 0
