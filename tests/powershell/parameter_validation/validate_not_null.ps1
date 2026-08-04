# vybe-test: powershell/parameter_validation/validate_not_null
function Test-Func { param([Parameter(Mandatory=$true)]$x); return $x }
if ((Test-Func -x 1) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
