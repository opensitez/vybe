# vybe-test: powershell/access_control/acl_path_type
$acl = Get-Acl $PWD
if (-not $acl.Path) {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
