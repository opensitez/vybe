# vybe-test: powershell/try_finally/function_finally
function Test-Func { try { $x = 1 } finally { $global:ran = $true } }
Test-Func
if ($ran) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
