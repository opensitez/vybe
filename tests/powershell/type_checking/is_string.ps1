# vybe-test: powershell/type_checking/is_string
if ('text' -is [string]) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
