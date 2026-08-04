# vybe-test: powershell/escape_sequences/single_quoted_literal
if ('A`nB' -eq 'A`nB') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
