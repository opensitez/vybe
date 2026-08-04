# vybe-test: powershell/scope_resolution/outer_scope
$x = 1
function Test-Func { if ($true) { $x = 2 } }
Test-Func
if ($x -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
