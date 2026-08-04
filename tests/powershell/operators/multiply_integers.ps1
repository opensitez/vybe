# vybe-test: powershell/operators/multiply_integers
$result = 6 * 7
if ($result -ne 42) {
    Write-Host "FAIL: expected 42, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
