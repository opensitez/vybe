# vybe-test: powershell/shift_operators/shift_left_assignment
$val = 4
$val = $val -shl 2
if ($val -ne 16) {
    Write-Host "FAIL: \$val -shl 2 expected 16, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
