# vybe-test: powershell/format_strings/format_boolean
if (("{0}" -f $true) -eq 'True') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
