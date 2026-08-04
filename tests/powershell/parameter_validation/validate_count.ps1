# vybe-test: powershell/parameter_validation/validate_count
function Test-Func { param([ValidateCount(1,2)]$x); return $x }
if ((Test-Func -x 1) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
