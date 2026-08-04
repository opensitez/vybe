# vybe-test: powershell/boolean_literals/or_false
if ($true -or $false) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
