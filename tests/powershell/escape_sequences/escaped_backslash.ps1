# vybe-test: powershell/escape_sequences/escaped_backslash
if ("Path\\File" -match '\\') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
