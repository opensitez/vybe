# vybe-test: powershell/string_methods/to_upper
if ('hello'.ToUpper() -eq 'HELLO') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
