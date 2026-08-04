# vybe-test: powershell/escape_sequences/backtick_escape
if ("Back``tick" -match 'Back`tick') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
