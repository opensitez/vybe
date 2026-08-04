# vybe-test: powershell/profile_loading/profile_loading_session
if (Test-Path $PROFILE) {
    . $PROFILE
}
Write-Host 'PASS'
exit 0
