# vybe-test: powershell/expression_precedence/precedence_add_multiply
if (1 + 2 * 3 -eq 7) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
