# vybe-test: powershell/shift_operators/shift_left_bitwise_and
$mask = 0xFF -band (1 -shl 4)
if ($mask -ne 16) {
    Write-Host "FAIL: 0xFF -band (1 -shl 4) expected 16, got $mask"
    exit 1
}
Write-Host "PASS"
exit 0
