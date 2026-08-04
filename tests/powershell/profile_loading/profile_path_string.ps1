# vybe-test: powershell/profile_loading/profile_path_string
if ($PROFILE -is [string]) {
    Write-Host 'PASS'
    exit 0
}
Write-Host "FAIL: expected string path"
exit 1
