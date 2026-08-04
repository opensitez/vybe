# vybe-test: powershell/operators/math_max
$result = [Math]::Max(10, 20)
if ($result -ne 20) {
    Write-Host "FAIL: expected 20, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
