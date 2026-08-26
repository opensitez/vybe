# vybe-test: powershell/bitwise_rotation_and_shifts/shift_modulo_32_masking
$x = 1
$shifted = $x -shl 33 # 33 % 32 = 1 -> 1 -shl 1 = 2
if ($shifted -ne 2) {
    Write-Host "FAIL: 1 -shl 33 expected 2, got $shifted"
    exit 1
}
Write-Host "PASS"
exit 0
