# vybe-test: powershell/boolean_literals/not_false
if (-not $false) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
