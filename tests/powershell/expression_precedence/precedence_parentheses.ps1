# vybe-test: powershell/expression_precedence/precedence_parentheses
if ((1 + 2) * 3 -eq 9) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
