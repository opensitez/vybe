# vybe-test: powershell/member_access/method_on_result
$result = 'hello'.ToUpper()
if ($result -ne 'HELLO') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
