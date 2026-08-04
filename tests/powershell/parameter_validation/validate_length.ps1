# vybe-test: powershell/parameter_validation/validate_length
function Test-Func { param([ValidateLength(1,3)]$x); return $x }
if ((Test-Func -x 'AB') -eq 'AB') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
