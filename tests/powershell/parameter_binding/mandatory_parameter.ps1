# vybe-test: powershell/parameter_binding/mandatory_parameter
function Test-Func { param([Parameter(Mandatory=$true)]$x); return $x }
if ((Test-Func -x 5) -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
