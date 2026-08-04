# vybe-test: powershell/method_calls/string_method
if ('hello'.ToUpper() -eq 'HELLO') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
