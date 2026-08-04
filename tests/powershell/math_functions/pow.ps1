# vybe-test: powershell/math_functions/pow
if ([math]::Pow(2,3) -eq 8) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
