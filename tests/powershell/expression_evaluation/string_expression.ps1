# vybe-test: powershell/expression_evaluation/string_expression
if ("Hello" + " World" -eq 'Hello World') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
