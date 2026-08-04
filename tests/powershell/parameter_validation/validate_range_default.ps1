# vybe-test: powershell/parameter_validation/validate_range_default
function Test-Func { param([ValidateRange(1,5)]$x = 3); return $x }
if ((Test-Func) -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
