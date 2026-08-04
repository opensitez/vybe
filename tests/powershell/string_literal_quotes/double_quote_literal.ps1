# vybe-test: powershell/string_literal_quotes/double_quote_literal
if ("Hello" -eq 'Hello') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
