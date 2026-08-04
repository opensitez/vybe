# vybe-test: powershell/array_methods/contains
if ((1,2,3) -contains 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
