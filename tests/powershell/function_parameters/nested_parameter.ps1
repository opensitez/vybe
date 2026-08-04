# vybe-test: powershell/function_parameters/nested_parameter
function Test-Func { param($x) return $x * 2 }
if ((Test-Func -x 4) -ne 8) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
