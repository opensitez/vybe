# vybe-test: powershell/pipeline_input/where_object
if ((1,2,3 | Where-Object { $_ -gt 1 }) -join ',' -eq '2,3') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
