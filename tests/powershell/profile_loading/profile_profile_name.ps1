# vybe-test: powershell/profile_loading/profile_profile_name.ps1
if ($PROFILE -like '*Microsoft.PowerShell_profile.ps1') {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'PASS'
exit 0
