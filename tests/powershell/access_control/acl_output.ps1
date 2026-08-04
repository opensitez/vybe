# vybe-test: powershell/access_control/acl_output
$acl = Get-Acl $PWD
if (-not $acl) {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
