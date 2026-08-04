# vybe-test: powershell/return_values/pipeline_return
$value = 1,2,3 | ForEach-Object { $_ }
if ($value.Count -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
