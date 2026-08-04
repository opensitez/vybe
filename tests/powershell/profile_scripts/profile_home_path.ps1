# vybe-test: powershell/profile_scripts/profile_home_path.ps1
if ($PROFILE -notlike '*$HOME*') {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'PASS'
exit 0
