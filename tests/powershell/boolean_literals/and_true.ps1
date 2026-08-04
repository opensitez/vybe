# vybe-test: powershell/boolean_literals/and_true
if ($true -and $true) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
