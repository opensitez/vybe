# vybe-test: powershell/shift_operators/shift_left_negative
$res = -1 -shl 1
if ($res -ne -2) {
    Write-Host "FAIL: -1 -shl 1 expected -2, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
