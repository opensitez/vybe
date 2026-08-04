# vybe-test: powershell/profile_scripts/profile_all_users.ps1
if (-not $PROFILE.AllUsersAllHosts) {
    Write-Host "FAIL: expected AllUsersAllHosts profile property"
    exit 1
}
Write-Host 'PASS'
exit 0
