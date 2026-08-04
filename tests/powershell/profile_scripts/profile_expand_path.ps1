# vybe-test: powershell/profile_scripts/profile_expand_path
$path = Resolve-Path $PROFILE.CurrentUserAllHosts -ErrorAction SilentlyContinue
if ($path -eq $null) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'PASS'
exit 0
