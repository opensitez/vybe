# vybe-test: powershell/shift_operators/shift_left_zero
$res = 42 -shl 0
if ($res -ne 42) {
    Write-Host "FAIL: 42 -shl 0 expected 42, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
