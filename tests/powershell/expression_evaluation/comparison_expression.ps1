# vybe-test: powershell/expression_evaluation/comparison_expression
if (3 -gt 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
