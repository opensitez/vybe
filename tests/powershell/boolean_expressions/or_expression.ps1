# vybe-test: powershell/boolean_expressions/or_expression
if ($true -or $false) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
