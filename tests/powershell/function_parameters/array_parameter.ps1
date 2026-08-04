# vybe-test: powershell/function_parameters/array_parameter
function Test-Func { param([int[]]$values) $values.Count }
if ((Test-Func -values 1,2,3) -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
