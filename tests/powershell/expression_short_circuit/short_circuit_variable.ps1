# vybe-test: powershell/expression_short_circuit/short_circuit_variable
$flag = $false
if ($false -and ($flag = $true)) { }
if (-not $flag) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
