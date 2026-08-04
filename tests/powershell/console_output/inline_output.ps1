# vybe-test: powershell/console_output/inline_output
$result = Write-Output ('a' + 'b')
if ($result -ne 'ab') {
    Write-Host "FAIL: expected ab"
    exit 1
}
Write-Host 'PASS'
exit 0
