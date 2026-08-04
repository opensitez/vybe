# vybe-test: powershell/variable_expansion/expand_variable_with_braces
$name = 'X'
if ("${name}y" -ne 'Xy') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
