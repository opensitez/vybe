# vybe-test: powershell/command_results/command_output_string
$value = $(Write-Output 'PASS')
if ($value -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
