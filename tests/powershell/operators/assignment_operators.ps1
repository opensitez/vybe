# vybe-test: powershell/operators/assignment_operators
$x = 10
$x += 5
if ($x -ne 15) {
    Write-Host "FAIL: expected 15 after +=, got $x"
    exit 1
}
$x -= 3
if ($x -ne 12) {
    Write-Host "FAIL: expected 12 after -=, got $x"
    exit 1
}
$x *= 2
if ($x -ne 24) {
    Write-Host "FAIL: expected 24 after *=, got $x"
    exit 1
}
$x /= 4
if ($x -ne 6) {
    Write-Host "FAIL: expected 6 after /=, got $x"
    exit 1
}
Write-Host "PASS"
exit 0
