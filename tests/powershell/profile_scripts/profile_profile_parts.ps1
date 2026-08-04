# vybe-test: powershell/profile_scripts/profile_profile_parts.ps1
if ($PROFILE  -eq $null) {
    Write-Host "FAIL: expected PROFILE value"
    exit 1
}
Write-Host 'PASS'
exit 0
