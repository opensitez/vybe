# vybe-test: powershell/escape_sequences/tab_escape
if ("A`tB" -match 'A') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
