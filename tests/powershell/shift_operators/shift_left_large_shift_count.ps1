# vybe-test: powershell/shift_operators/shift_left_large_shift_count
$res = 1 -shl 31
if ($res -ne [int]::MinValue) {
    Write-Host "FAIL: 1 -shl 31 expected [int]::MinValue (-2147483648), got $res"
    exit 1
}
Write-Host "PASS"
exit 0
