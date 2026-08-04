# vybe-test: powershell/format_strings/format_multiple
if (("{0} {1}" -f 'A','B') -eq 'A B') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
