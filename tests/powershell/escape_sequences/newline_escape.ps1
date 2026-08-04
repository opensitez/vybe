# vybe-test: powershell/escape_sequences/newline_escape
if ("Line1`nLine2" -match 'Line2') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
