# vybe-test: powershell/escape_sequences/quote_escape
if ("He said \"Hi\"" -match 'Hi') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
