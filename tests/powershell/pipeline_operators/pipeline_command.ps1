# vybe-test: powershell/pipeline_operators/pipeline_command
if (('a','b' | ForEach-Object { $_.ToUpper() }) -join ',' -eq 'A,B') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
