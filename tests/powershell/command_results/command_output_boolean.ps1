# vybe-test: powershell/command_results/command_output_boolean
$value = (1 -eq 1)
if ($value) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
