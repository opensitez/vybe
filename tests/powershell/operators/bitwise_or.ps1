# vybe-test: powershell/operators/bitwise_or
$result = 12 -bor 10
if ($result -ne 14) {
    Write-Host "FAIL: expected 14, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
