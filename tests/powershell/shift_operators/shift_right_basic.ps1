# vybe-test: powershell/shift_operators/shift_right_basic
$res = 16 -shr 2
if ($res -ne 4) {
    Write-Host "FAIL: 16 -shr 2 expected 4, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
