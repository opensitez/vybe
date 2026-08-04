# vybe-test: powershell/function_scope/global_variable
function Test-Func { $global:x = 2 }
Test-Func
if ($x -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
