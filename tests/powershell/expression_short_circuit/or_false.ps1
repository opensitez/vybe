# vybe-test: powershell/expression_short_circuit/or_false
if ($false -or $false) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
