# vybe-test: powershell/string_methods/split
if (('a,b'.Split(',')).Count -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
