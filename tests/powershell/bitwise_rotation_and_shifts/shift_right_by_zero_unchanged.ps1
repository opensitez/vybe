# vybe-test: powershell/bitwise_rotation_and_shifts/shift_right_by_zero_unchanged
$x = 42
$shifted = $x -shr 0
if ($shifted -ne 42) {
    Write-Host "FAIL: Shift right by 0 failed"
    exit 1
}
Write-Host "PASS"
exit 0
