# vybe-test: powershell/type_unsigned_integers/bitwise_shift_right_logical_uint32
[uint32]$x = 2147483648
$shifted = $x -shr 1
if ($shifted -ne 1073741824) {
    Write-Host "FAIL: uint32 right shift mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
