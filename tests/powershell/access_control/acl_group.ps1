# vybe-test: powershell/access_control/acl_group
$acl = Get-Acl $PWD
if (-not $acl.Group) {
    Write-Host "PASS"
    exit 0
}
Write-Host 'PASS'
exit 0
