# vybe-test: powershell/parameter_validation/validate_set
function Test-Func { param([ValidateSet('A','B')]$x); return $x }
if ((Test-Func -x 'A') -eq 'A') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
