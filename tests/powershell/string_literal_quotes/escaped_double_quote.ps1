# vybe-test: powershell/string_literal_quotes/escaped_double_quote
if ("He said \"Hi\"" -match 'Hi') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
