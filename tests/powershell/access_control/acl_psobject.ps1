# vybe-test: powershell/access_control/acl_psobject
$acl = Get-Acl $PWD
if (-not ($acl -is [object])) {
    Write-Host "FAIL: expected object"
    exit 1
}
Write-Host 'PASS'
exit 0
