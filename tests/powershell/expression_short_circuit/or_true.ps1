# vybe-test: powershell/expression_short_circuit/or_true
$called = $false
if ($true -or ($called = $true)) { }
if (-not $called) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
