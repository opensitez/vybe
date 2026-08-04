# vybe-test: powershell/method_calls/method_on_result
if ((1..3).Count.ToString() -eq '3') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
