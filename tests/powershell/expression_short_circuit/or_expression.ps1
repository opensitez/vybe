# vybe-test: powershell/expression_short_circuit/or_expression
if (($false -or $true) -or $false) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
