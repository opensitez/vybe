# vybe-test: powershell/profile_scripts/profile_profile_type.ps1
if ($PROFILE -is [string]) {
    Write-Host 'PASS'
    exit 0
}
Write-Host "FAIL: expected PROFILE string"
exit 1
