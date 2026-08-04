# vybe-test: powershell/pipeline_operators/pipeline_sort
if ((3,1,2 | Sort-Object) -join ',' -eq '1,2,3') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
