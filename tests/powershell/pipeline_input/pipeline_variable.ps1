# vybe-test: powershell/pipeline_input/pipeline_variable
if ((1,2,3 | ForEach-Object { $PSItem }) -join ',' -eq '1,2,3') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
