# vybe-test: powershell/function_parameters/parameter_alias
function Test-Func { param([Alias('n')]$x) $x }
if ((Test-Func -n 5) -ne 5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
