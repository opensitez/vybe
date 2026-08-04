# vybe-test: powershell/logical_operators/short_circuit_or
$called = $false
if ($true -or ($called = $true)) {
}
if ($called -eq $false) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
