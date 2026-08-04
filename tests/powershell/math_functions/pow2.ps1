# vybe-test: powershell/math_functions/pow2
if ([math]::Pow(2,5) -eq 32) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
