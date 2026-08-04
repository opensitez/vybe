# vybe-test: powershell/operators/math_min
$result = [Math]::Min(10, 20)
if ($result -ne 10) {
    Write-Host "FAIL: expected 10, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
