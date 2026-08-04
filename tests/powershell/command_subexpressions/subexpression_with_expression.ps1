# vybe-test: powershell/command_subexpressions/subexpression_with_expression
if ("$((2 * 3))" -eq '6') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
