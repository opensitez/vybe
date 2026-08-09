# vybe-test: powershell/shift_operators/shift_right_sign_extension
$res = -8 -shr 2
if ($res -ne -2) {
    Write-Host "FAIL: -8 -shr 2 expected -2, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
