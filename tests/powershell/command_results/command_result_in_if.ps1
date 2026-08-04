# vybe-test: powershell/command_results/command_result_in_if
if ((Write-Output 'PASS') -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
