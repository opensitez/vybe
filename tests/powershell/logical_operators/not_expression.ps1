# vybe-test: powershell/logical_operators/not_expression
if (-not ($false -or $false)) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
