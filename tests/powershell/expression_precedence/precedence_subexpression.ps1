# vybe-test: powershell/expression_precedence/precedence_subexpression
if ($(1 + 1) -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
