# vybe-test: powershell/command_results/command_output_array
$value = (1,2,3)
if ($value.Count -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
