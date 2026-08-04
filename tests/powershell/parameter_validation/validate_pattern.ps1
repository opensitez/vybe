# vybe-test: powershell/parameter_validation/validate_pattern
function Test-Func { param([ValidatePattern('^A') ]$x); return $x }
if ((Test-Func -x 'A1') -eq 'A1') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
