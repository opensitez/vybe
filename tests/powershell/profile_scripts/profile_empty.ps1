# vybe-test: powershell/profile_scripts/profile_empty.ps1
if ($PROFILE -eq '') {
    Write-Host "FAIL: expected non-empty PROFILE"
    exit 1
}
Write-Host 'PASS'
exit 0
