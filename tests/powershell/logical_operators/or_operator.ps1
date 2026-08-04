# vybe-test: powershell/logical_operators/or_operator
if ($false -or $true) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
