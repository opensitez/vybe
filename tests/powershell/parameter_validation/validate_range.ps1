# vybe-test: powershell/parameter_validation/validate_range
function Test-Func { param([ValidateRange(1,10)]$x); return $x }
if ((Test-Func -x 5) -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
