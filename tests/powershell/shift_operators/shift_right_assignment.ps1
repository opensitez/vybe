# vybe-test: powershell/shift_operators/shift_right_assignment
$val = 64
$val = $val -shr 3
if ($val -ne 8) {
    Write-Host "FAIL: \$val -shr 3 expected 8, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
