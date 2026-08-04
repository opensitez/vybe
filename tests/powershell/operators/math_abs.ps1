# vybe-test: powershell/operators/math_abs
$result = [Math]::Abs(-42)
if ($result -ne 42) {
    Write-Host "FAIL: expected 42, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
