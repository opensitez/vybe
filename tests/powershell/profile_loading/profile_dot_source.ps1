# vybe-test: powershell/profile_loading/profile_dot_source
if (Test-Path $PROFILE) {
    . $PROFILE
}
Write-Host 'PASS'
exit 0
