# vybe-test: powershell/function_parameters/default_parameter
function Test-Func { param($x = 10) $x }
if ((Test-Func) -ne 10) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
