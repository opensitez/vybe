# vybe-test: powershell/operators/bitwise_xor
$result = 12 -bxor 10
if ($result -ne 6) {
    Write-Host "FAIL: expected 6, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
