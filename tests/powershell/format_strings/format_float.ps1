# vybe-test: powershell/format_strings/format_float
if (("{0:F1}" -f 1.23) -eq '1.2') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
