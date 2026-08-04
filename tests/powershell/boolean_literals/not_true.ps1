# vybe-test: powershell/boolean_literals/not_true
if (-not $true -eq $false) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
