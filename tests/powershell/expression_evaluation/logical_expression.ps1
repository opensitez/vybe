# vybe-test: powershell/expression_evaluation/logical_expression
if ((1 -eq 1) -and (2 -eq 2)) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
