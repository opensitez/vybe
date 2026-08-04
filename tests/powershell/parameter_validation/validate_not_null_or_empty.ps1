# vybe-test: powershell/parameter_validation/validate_not_null_or_empty
function Test-Func { param([ValidateNotNullOrEmpty()]$x); return $x }
if ((Test-Func -x 'PASS') -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
