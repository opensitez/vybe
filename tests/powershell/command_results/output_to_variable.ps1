# vybe-test: powershell/command_results/output_to_variable
$value = Write-Output 'PASS'
if ($value -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
