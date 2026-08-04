# vybe-test: powershell/scope_resolution/script_scope
function Test-Func { $script:x = 3 }
Test-Func
if ($script:x -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
