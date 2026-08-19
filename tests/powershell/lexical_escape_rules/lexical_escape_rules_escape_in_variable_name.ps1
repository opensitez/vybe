# vybe-test: powershell/lexical_escape_rules/escape_in_variable_name
${na me} = 12
if (${na me} -ne 12) {
    Write-Host 'FAIL: escaped variable name with backtick-space not assigned correctly'
    exit 1
}

Write-Host 'PASS'
exit 0
