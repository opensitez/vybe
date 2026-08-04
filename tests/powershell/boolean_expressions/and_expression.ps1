# vybe-test: powershell/boolean_expressions/and_expression
if ($true -and $true) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
