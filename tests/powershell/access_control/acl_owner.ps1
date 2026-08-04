# vybe-test: powershell/access_control/acl_owner
$acl = Get-Acl $PWD
if (-not $acl.Owner) {
    Write-Host "FAIL: expected owner"
    exit 1
}
Write-Host 'PASS'
exit 0
