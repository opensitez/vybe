# vybe-test: powershell/format_strings/date_format
if ((Get-Date).ToString('yyyy') -match '2026') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
