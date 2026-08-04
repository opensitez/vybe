# vybe-test: powershell/profile_loading/profile_loading_path_type.ps1
if ($PROFILE -is [string]) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL: expected string profile'
exit 1
