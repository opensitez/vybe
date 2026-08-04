# vybe-test: powershell/expressions/string_expression
if ("Hello" -eq 'Hello') {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
