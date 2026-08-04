# vybe-test: powershell/operators/modulo_operator
$result = 17 % 5
if ($result -ne 2) {
    Write-Host "FAIL: expected 2, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
