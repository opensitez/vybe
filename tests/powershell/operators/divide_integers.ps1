# vybe-test: powershell/operators/divide_integers
$result = 20 / 4
if ($result -ne 5) {
    Write-Host "FAIL: expected 5, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
