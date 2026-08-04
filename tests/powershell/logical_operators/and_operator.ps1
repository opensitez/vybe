# vybe-test: powershell/logical_operators/and_operator
if ($true -and $true) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
