# vybe-test: powershell/types/type_casting_int
$x = [int]"42"
if ($x -ne 42) {
    Write-Host "FAIL: expected 42, got $x"
    exit 1
}
Write-Host "PASS"
exit 0
