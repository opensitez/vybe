# vybe-test: powershell/operators/math_sqrt
$result = [Math]::Sqrt(16)
if ($result -ne 4) {
    Write-Host "FAIL: expected 4, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
