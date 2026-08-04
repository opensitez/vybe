# vybe-test: powershell/parameter_validation/validate_set_named
function Test-Func { param([ValidateSet('A','B')]$x); return $x }
if ((Test-Func -x 'B') -eq 'B') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
