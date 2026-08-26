# vybe-test: powershell/bitwise_rotation_and_shifts/bitwise_shift_left_preserves_sign_or_wraps
[int32]$val = 1
$res = $val -shl 31
if ($res -ne [int32]::MinValue) {
    Write-Host "FAIL: Shift left into sign bit failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
