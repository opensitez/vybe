# vybe-test: powershell/pipeline_input/simple_pipeline
if ((1,2,3 | ForEach-Object { $_ }) -join ',' -eq '1,2,3') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
