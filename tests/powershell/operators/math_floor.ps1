# vybe-test: powershell/operators/math_floor
$result = [Math]::Floor(3.7)
if ($result -ne 3) {
    Write-Host "FAIL: expected 3, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
