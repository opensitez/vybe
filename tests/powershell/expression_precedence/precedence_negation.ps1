# vybe-test: powershell/expression_precedence/precedence_negation
if ((-1) -lt 0) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
