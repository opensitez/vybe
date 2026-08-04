# vybe-test: powershell/operators/bitwise_and
$result = 12 -band 10
if ($result -ne 8) {
    Write-Host "FAIL: expected 8, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
