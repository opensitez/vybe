# vybe-test: powershell/escape_sequences/dollar_escape
if ("`$value" -match '\$value') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
