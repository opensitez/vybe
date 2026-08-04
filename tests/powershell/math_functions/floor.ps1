# vybe-test: powershell/math_functions/floor
if ([math]::Floor(1.9) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
