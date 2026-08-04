# vybe-test: powershell/string_literal_quotes/quotes_within_double
if ("She said 'Hi'" -match "Hi") { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
