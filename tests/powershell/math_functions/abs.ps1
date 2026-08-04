# vybe-test: powershell/math_functions/abs
if ([math]::Abs(-5) -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
