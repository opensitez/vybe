# vybe-test: powershell/function_scope/loop_scope
function Test-Func { for ($i=0; $i -lt 1; $i++) { $x = 2 } }
Test-Func
if ($x -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
