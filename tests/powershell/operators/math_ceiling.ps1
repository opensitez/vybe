# vybe-test: powershell/operators/math_ceiling
$result = [Math]::Ceiling(3.2)
if ($result -ne 4) {
    Write-Host "FAIL: expected 4, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
