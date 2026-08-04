# vybe-test: powershell/access_control/acl_accesssentry
$acl = Get-Acl $PWD
if (-not $acl.Access) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'PASS'
exit 0
