# vybe-test: powershell/operators/increment_operator
$x = 5
$x++
if ($x -ne 6) {
    Write-Host "FAIL: expected 6, got $x"
    exit 1
}
Write-Host "PASS"
exit 0
