# vybe-test: powershell/boolean_literals/comparison_true
if (1 -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
