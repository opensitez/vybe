# vybe-test: powershell/string_methods/replace
if ('hello'.Replace('h','j') -eq 'jello') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
