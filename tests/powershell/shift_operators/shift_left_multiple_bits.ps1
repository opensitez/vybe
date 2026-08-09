# vybe-test: powershell/shift_operators/shift_left_multiple_bits
$res = 5 -shl 4
if ($res -ne 80) {
    Write-Host "FAIL: 5 -shl 4 expected 80, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
