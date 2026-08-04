# vybe-test: powershell/scope_resolution/variable_shadowing
$x = 1
function Test-Func { param($x); return $x }
if ((Test-Func 2) -eq 2 -and $x -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
