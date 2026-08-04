# vybe-test: powershell/expression_short_circuit/and_true
if ($true -and $true) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
