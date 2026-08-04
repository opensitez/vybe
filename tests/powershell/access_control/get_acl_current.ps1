# vybe-test: powershell/access_control/get_acl_current
$path = $PWD
$acl = Get-Acl -Path $path
if (-not $acl) {
    Write-Host "FAIL: expected ACL"
    exit 1
}
Write-Host 'PASS'
exit 0
