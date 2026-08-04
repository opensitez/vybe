# vybe-test: powershell/string_methods/trim
if ('  hi  '.Trim() -eq 'hi') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
