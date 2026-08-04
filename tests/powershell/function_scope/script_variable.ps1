# vybe-test: powershell/function_scope/script_variable
function Test-Func { $script:x = 3 }
Test-Func
if ($script:x -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
