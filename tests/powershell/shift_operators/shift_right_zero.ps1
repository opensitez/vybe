# vybe-test: powershell/shift_operators/shift_right_zero
$res = 99 -shr 0
if ($res -ne 99) {
    Write-Host "FAIL: 99 -shr 0 expected 99, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
