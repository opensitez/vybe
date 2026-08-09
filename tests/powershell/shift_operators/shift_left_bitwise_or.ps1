# vybe-test: powershell/shift_operators/shift_left_bitwise_or
$res = (1 -shl 2) -bor (1 -shl 3)
if ($res -ne 12) {
    Write-Host "FAIL: (1 -shl 2) -bor (1 -shl 3) expected 12, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
