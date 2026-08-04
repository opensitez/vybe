# vybe-test: powershell/expression_evaluation/pipeline_expression
if ((1,2,3 | Measure-Object).Count -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
