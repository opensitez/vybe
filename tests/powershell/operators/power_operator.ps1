# vybe-test: powershell/operators/power_operator
$result = [Math]::Pow(2, 3)
if ($result -ne 8) {
    Write-Host "FAIL: expected 8, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
