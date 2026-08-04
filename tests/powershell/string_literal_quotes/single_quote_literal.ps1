# vybe-test: powershell/string_literal_quotes/single_quote_literal
if ('Hello' -eq 'Hello') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
