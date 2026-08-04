# vybe-test: powershell/expressions/function_expression
if ((Get-Date).Year -ge 2026) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
