# vybe-test: powershell/pipeline_input/pipeline_expression
if ((1..3 | Where-Object { $_ -lt 3 }).Count -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
