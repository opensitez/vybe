# vybe-test: powershell/format_strings/format_padding
if (("{0,4}" -f 'x').Trim() -eq 'x') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
