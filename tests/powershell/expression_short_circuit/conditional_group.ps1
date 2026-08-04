# vybe-test: powershell/expression_short_circuit/conditional_group
if ((1 -eq 1) -and ($true -or $false)) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
