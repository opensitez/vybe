# vybe-test: powershell/expression_short_circuit/and_expression
if (($true -and $true) -and $true) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
