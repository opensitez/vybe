# vybe-test: powershell/bitwise_rotation_and_shifts/shift_left_multiple_bits
$x = 1
$shifted = $x -shl 4
if ($shifted -ne 16) {
    Write-Host "FAIL: 1 -shl 4 expected 16, got $shifted"
    exit 1
}
Write-Host "PASS"
exit 0
