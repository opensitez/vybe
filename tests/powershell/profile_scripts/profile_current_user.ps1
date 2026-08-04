# vybe-test: powershell/profile_scripts/profile_current_user
if (-not $PROFILE.CurrentUserAllHosts) {
    Write-Host "FAIL: expected CurrentUserAllHosts profile property"
    exit 1
}
Write-Host 'PASS'
exit 0
