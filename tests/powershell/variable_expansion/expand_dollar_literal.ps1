# vybe-test: powershell/variable_expansion/expand_dollar_literal
if ("`$x" -ne '$x') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
