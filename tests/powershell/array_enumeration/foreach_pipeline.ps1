# vybe-test: powershell/array_enumeration/foreach_pipeline
$result = 1,2,3 | ForEach-Object { $_ * 2 }
if ($result -join ',' -eq '2,4,6') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
