# vybe-test: powershell/escape_sequences/carriage_return_escape
if ("A`rB" -match 'A') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
