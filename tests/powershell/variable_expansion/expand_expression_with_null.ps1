# vybe-test: powershell/variable_expansion/expand_expression_with_null
if ("$($null)" -ne '') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
