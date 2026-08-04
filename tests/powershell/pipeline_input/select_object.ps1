# vybe-test: powershell/pipeline_input/select_object
if ((1,2,3 | Select-Object -First 2) -join ',' -eq '1,2') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
