# vybe-test: powershell/profile_scripts/profile_path_exists
if (-not $PROFILE) {
    Write-Host "FAIL: expected PROFILE defined"
    exit 1
}
Write-Host 'PASS'
exit 0
