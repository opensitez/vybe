# vybe-test: powershell/pipeline_operators/pipeline_expression
if ((1,2,3 | ForEach-Object { $_ * 2 }) -join ',' -eq '2,4,6') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
