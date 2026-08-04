# vybe-test: powershell/command_results/command_chain_result
$value = Write-Output 'PASS' | ForEach-Object { $_ }
if ($value -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
