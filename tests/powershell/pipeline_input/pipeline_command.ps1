# vybe-test: powershell/pipeline_input/pipeline_command
if ((1,2,3 | Sort-Object -Descending) -join ',' -eq '3,2,1') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
