# vybe-test: powershell/profile_scripts/profile_all_hosts.ps1
if (-not $PROFILE.AllUsersCurrentHost) {
    Write-Host "FAIL: expected AllUsersCurrentHost profile property"
    exit 1
}
Write-Host 'PASS'
exit 0
