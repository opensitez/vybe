# vybe-test: powershell/expression_short_circuit/and_false
$called = $false
if ($false -and ($called = $true)) { }
if (-not $called) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
