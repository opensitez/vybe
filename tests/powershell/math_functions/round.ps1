# vybe-test: powershell/math_functions/round
if ([math]::Round(1.5) -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
