# vybe-test: powershell/boolean_expressions/boolean_pipeline
if ((1,2,3 | Where-Object { $_ -gt 0 }) -ne $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
