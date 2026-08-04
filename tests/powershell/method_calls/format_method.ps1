# vybe-test: powershell/method_calls/format_method
if ([string]::Format('{0} {1}','A','B') -eq 'A B') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
