# vybe-test: powershell/shift_operators/shift_right_uint32_type
$u = [uint32]100
$res = $u -shr 1
if ($res -ne 50) {
    Write-Host "FAIL: [uint32]100 -shr 1 expected 50, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
