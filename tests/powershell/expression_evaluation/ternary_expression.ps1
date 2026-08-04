# vybe-test: powershell/expression_evaluation/ternary_expression
if ((1 -eq 1) ? 'PASS' : 'FAIL' -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
