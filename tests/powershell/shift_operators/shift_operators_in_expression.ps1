# vybe-test: powershell/shift_operators/shift_operators_in_expression
$res = (1 -shl 10) / 1024
if ($res -ne 1) {
    Write-Host "FAIL: (1 -shl 10) / 1024 expected 1, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
