# vybe-test: powershell/escape_sequences/escaped_variable_in_string
if ("`$x" -match '\$x') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
