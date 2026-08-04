# vybe-test: powershell/try_finally/returned_value
function Test-Func { try { return 'PASS' } finally { $global:ran = $true } }
if ((Test-Func) -eq 'PASS' -and $ran) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
