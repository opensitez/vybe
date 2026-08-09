# vybe-test: powershell/shift_operators/shift_left_overflow_to_int64
$res = [int64]1 -shl 35
if ($res -ne 34359738368) {
    Write-Host "FAIL: [int64]1 -shl 35 expected 34359738368, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
