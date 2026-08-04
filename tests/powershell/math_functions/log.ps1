# vybe-test: powershell/math_functions/log
if ([math]::Log(1) -eq 0) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
