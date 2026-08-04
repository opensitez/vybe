# vybe-test: powershell/format_strings/padded_string
if (("{0,5}" -f 'x').Trim() -eq 'x') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
