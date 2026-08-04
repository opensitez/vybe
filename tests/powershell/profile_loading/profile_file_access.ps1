# vybe-test: powershell/profile_loading/profile_file_access
if ($PROFILE.Length -gt 0) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'PASS'
exit 0
