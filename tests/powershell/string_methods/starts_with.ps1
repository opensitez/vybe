# vybe-test: powershell/string_methods/starts_with
if ('hello'.StartsWith('he')) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
