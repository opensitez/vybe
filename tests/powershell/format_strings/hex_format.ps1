# vybe-test: powershell/format_strings/hex_format
if (("{0:X}" -f 255) -eq 'FF') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
