# vybe-test: powershell/string_methods/substring
if ('hello'.Substring(1,3) -eq 'ell') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
