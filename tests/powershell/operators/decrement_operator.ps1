# vybe-test: powershell/operators/decrement_operator
$x = 10
$x--
if ($x -ne 9) {
    Write-Host "FAIL: expected 9, got $x"
    exit 1
}
Write-Host "PASS"
exit 0
