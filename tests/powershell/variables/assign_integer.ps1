# vybe-test: powershell/variables/assign_integer
$x = 42
if ($x -ne 42) {
    Write-Host "FAIL: expected 42, got $x"
    exit 1
}
Write-Host "PASS"
exit 0
