# vybe-test: powershell/array_methods/join
if ((1,2,3) -join ',' -eq '1,2,3') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
