# vybe-test: powershell/command_results/multiple_outputs
$value = ,(Write-Output 'A'; Write-Output 'B')
if ($value.Count -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
