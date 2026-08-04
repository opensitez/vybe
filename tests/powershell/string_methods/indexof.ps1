# vybe-test: powershell/string_methods/indexof
if ('hello'.IndexOf('e') -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
