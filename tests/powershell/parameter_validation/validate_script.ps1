# vybe-test: powershell/parameter_validation/validate_script
function Test-Func { param([ValidateScript({ $_ -gt 0 })]$x); return $x }
if ((Test-Func -x 1) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
