# vybe-test: powershell/function_parameters/switch_parameter
function Test-Func { param([switch]$Flag) if ($Flag) { 'yes' } else { 'no' } }
if ((Test-Func -Flag) -ne 'yes') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
