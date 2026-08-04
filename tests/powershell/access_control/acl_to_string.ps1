# vybe-test: powershell/access_control/acl_to_string
$acl = Get-Acl $PWD
if ($acl.ToString() -eq $null) {
    Write-Host "FAIL: expected string"
    exit 1
}
Write-Host 'PASS'
exit 0
