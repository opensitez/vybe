# vybe-test: powershell/math_functions/sqrt
if ([math]::Sqrt(9) -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
