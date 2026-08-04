# vybe-test: powershell/profile_scripts/profile_file_name.ps1
if ($PROFILE -notlike '*.ps1') {
    Write-Host "FAIL: expected profile file extension"
    exit 1
}
Write-Host 'PASS'
exit 0
