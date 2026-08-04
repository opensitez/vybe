# vybe-test: powershell/function_parameters/mandatory_parameter
function Test-Func { param([Parameter(Mandatory=$true)] $x) $x }
$result = Test-Func -x 5
if ($result -ne 5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
