# vybe-test: powershell/command_results/command_output_property
$value = (Get-Date).Year
if ($value -gt 2000) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
