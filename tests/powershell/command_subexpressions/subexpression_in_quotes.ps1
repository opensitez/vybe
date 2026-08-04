# vybe-test: powershell/command_subexpressions/subexpression_in_quotes
if ("A$((1 + 1))B" -eq 'A2B') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
