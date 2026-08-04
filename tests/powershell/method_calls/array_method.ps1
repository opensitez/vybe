# vybe-test: powershell/method_calls/array_method
if ((1,2,3).Count -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
