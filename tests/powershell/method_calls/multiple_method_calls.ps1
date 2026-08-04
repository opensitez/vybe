# vybe-test: powershell/method_calls/multiple_method_calls
if ('hello'.ToUpper().Substring(0,2) -eq 'HE') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
