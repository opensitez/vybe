# vybe-test: powershell/function_parameters/parameter_validation
function Test-Func { param([ValidateRange(1,3)]$x) return $x }
if ((Test-Func -x 2) -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
