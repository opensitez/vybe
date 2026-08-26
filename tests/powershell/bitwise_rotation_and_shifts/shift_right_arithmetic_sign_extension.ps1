# vybe-test: powershell/bitwise_rotation_and_shifts/shift_right_arithmetic_sign_extension
$x = -8
$shifted = $x -shr 1 # arithmetic shift preserves sign
if ($shifted -ne -4) {
    Write-Host "FAIL: Arithmetic right shift -8 -shr 1 expected -4, got $shifted"
    exit 1
}
Write-Host "PASS"
exit 0
