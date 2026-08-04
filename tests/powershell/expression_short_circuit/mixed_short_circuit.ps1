# vybe-test: powershell/expression_short_circuit/mixed_short_circuit
$called = $false
if (($false -and ($called = $true)) -or $true) { }
if (-not $called) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
