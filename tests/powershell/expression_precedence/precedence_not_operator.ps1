# vybe-test: powershell/expression_precedence/precedence_not_operator
if (-not $false) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
