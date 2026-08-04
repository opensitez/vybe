# vybe-test: powershell/profile_loading/profile_exists_or_not
if ($PROFILE -ne $null) {
    Write-Host 'PASS'
    exit 0
}
Write-Host "FAIL: expected profile variable"
exit 1
