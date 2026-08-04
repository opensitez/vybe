# vybe-test: powershell/pipeline_input/foreach_object
if ((1,2 | ForEach-Object { $_ * 2 }) -join ',' -eq '2,4') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
