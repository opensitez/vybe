# vybe-test: powershell/shift_operators/shift_left_basic
$res = 1 -shl 3
if ($res -ne 8) {
    Write-Host "FAIL: 1 -shl 3 expected 8, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
