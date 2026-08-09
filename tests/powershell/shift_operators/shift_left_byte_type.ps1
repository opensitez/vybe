# vybe-test: powershell/shift_operators/shift_left_byte_type
$b = [byte]2
$res = $b -shl 3
if ($res -ne 16) {
    Write-Host "FAIL: [byte]2 -shl 3 expected 16, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
