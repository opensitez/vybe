# vybe-test: powershell/pipeline_operators/pipeline_to_variable
$result = 1,2,3 | Where-Object { $_ -gt 1 }
if ($result -join ',' -eq '2,3') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
