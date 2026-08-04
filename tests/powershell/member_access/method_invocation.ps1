# vybe-test: powershell/member_access/method_invocation
$str = 'hello'
if ($str.Length -ne 5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
