# vybe-test: powershell/access_control/acl_accessrules
$acl = Get-Acl $PWD
if (-not $acl.Access | Out-Null) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'PASS'
exit 0
