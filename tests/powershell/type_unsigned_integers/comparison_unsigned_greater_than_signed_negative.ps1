# vybe-test: powershell/type_unsigned_integers/comparison_unsigned_greater_than_signed_negative
[uint32]$u = 10
$neg = -5
if (-not ($u -gt $neg)) {
    Write-Host "FAIL: unsigned 10 should be greater than -5"
    exit 1
}
Write-Host "PASS"
exit 0
