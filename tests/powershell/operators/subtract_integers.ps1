# vybe-test: powershell/operators/subtract_integers
$result = 10 - 4
if ($result -ne 6) {
    Write-Host "FAIL: expected 6, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
