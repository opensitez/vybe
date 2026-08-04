# vybe-test: powershell/math_functions/max
if ([math]::Max(1,2) -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
