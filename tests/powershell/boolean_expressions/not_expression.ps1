# vybe-test: powershell/boolean_expressions/not_expression
if (-not $false) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
