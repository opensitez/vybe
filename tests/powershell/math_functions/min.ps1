# vybe-test: powershell/math_functions/min
if ([math]::Min(1,2) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
