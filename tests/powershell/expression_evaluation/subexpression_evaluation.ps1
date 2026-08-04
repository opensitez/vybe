# vybe-test: powershell/expression_evaluation/subexpression_evaluation
if ($((1 + 1)) -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
