# vybe-test: powershell/variable_expansion/expand_expression_in_string
if ("1+1=$([int]1+1)" -ne '1+1=2') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
