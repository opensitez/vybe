# vybe-test: powershell/string_literal_quotes/subexpression
if ("$(1 + 1)" -eq '2') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
