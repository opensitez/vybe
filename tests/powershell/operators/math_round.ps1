# vybe-test: powershell/operators/math_round
$result = [Math]::Round(3.6)
if ($result -ne 4) {
    Write-Host "FAIL: expected 4, got $result"
    exit 1
}
$result2 = [Math]::Round(3.4)
if ($result2 -ne 3) {
    Write-Host "FAIL: expected 3, got $result2"
    exit 1
}
Write-Host "PASS"
exit 0
