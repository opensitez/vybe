# vybe-test: powershell/access_control/get_acl_file
$path = $PWD
$acl = Get-Acl $path
if (-not $acl.Path) {
    Write-Host "FAIL: expected ACL path"
    exit 1
}
Write-Host 'PASS'
exit 0
