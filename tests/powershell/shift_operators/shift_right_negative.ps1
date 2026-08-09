# vybe-test: powershell/shift_operators/shift_right_negative
$res = -16 -shr 3
if ($res -ne -2) {
    Write-Host "FAIL: -16 -shr 3 expected -2, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
