# vybe-test: powershell/boolean_literals/boolean_expression
if (($true -and $false) -or $true) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
