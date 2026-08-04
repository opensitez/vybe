# vybe-test: powershell/escape_sequences/subexpression_escape
if ("`$(1 + 1)" -match '\$\(') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
