# vybe-test: powershell/array_methods/count
if ((1,2,3).Count -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
