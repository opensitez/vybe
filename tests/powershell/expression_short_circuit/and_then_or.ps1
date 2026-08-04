# vybe-test: powershell/expression_short_circuit/and_then_or
if (($true -and $false) -or $true) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
