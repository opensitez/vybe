# vybe-test: powershell/variables/reassign_variable
$x = 10
$x = 20
if ($x -ne 20) {
    Write-Host "FAIL: expected 20, got $x"
    exit 1
}
Write-Host "PASS"
exit 0
