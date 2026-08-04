# vybe-test: powershell/expression_evaluation/string_concat_expression
if (("A" + "B") -eq 'AB') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
