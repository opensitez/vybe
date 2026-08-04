# vybe-test: powershell/boolean_literals/false_literal
if (-not $false) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
