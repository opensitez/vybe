# vybe-test: powershell/string_methods/ends_with
if ('hello'.EndsWith('lo')) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
