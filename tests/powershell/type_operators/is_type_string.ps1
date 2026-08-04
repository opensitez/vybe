# vybe-test: powershell/type_operators/is_type_string
if ('text' -is [string]) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
