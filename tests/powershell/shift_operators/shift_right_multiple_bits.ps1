# vybe-test: powershell/shift_operators/shift_right_multiple_bits
$res = 256 -shr 4
if ($res -ne 16) {
    Write-Host "FAIL: 256 -shr 4 expected 16, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
