# vybe-test: powershell/function_parameters/positional_parameters
function Test-Func { param($x, $y) return $x + $y }
if ((Test-Func 1 2) -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
