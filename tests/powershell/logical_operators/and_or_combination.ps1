# vybe-test: powershell/logical_operators/and_or_combination
if (($true -and $true) -or $false) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
