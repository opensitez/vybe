# vybe-test: powershell/function_parameters/named_parameters
function Test-Func { param($x, $y) return $x + $y }
if ((Test-Func -y 2 -x 1) -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
