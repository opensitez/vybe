# vybe-test: powershell/pipeline_input/pipeline_assignment
$arr = 1,2,3 | Where-Object { $_ -gt 1 }
if ($arr -join ',' -eq '2,3') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
