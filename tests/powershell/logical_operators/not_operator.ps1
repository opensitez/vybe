# vybe-test: powershell/logical_operators/not_operator
if (-not $false) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
