# vybe-test: powershell/method_calls/static_method
if ([math]::Sqrt(4) -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
