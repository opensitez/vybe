# vybe-test: powershell/pipeline_operators/pipeline_group
if ((1,2,3 | Group-Object).Count -ge 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
