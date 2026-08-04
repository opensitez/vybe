# vybe-test: powershell/type_checking/is_object
if ('text' -is [object]) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
