# vybe-test: powershell/pipeline_input/multiple_stages
if ((1,2 | ForEach-Object { $_ + 1 } | Where-Object { $_ -gt 1 }) -join ',' -eq '2,3') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
