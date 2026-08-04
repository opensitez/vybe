# vybe-test: powershell/math_functions/ceiling
if ([math]::Ceiling(1.1) -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
