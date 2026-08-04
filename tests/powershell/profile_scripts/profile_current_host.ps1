# vybe-test: powershell/profile_scripts/profile_current_host.ps1
if (-not $PROFILE.CurrentUserCurrentHost) {
    Write-Host "FAIL: expected CurrentUserCurrentHost profile property"
    exit 1
}
Write-Host 'PASS'
exit 0
