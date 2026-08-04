# vybe-test: powershell/boolean_literals/boolean_variable
$val = $true
if ($val) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
