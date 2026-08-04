# vybe-test: powershell/format_strings/format_number
if (("{0:N0}" -f 1000) -eq '1,000') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
