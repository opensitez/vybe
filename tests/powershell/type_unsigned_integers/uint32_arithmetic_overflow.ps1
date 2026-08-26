# vybe-test: powershell/type_unsigned_integers/uint32_arithmetic_overflow
[uint32]$u1 = 4000000000
[uint32]$u2 = 1000000000
$sum = $u1 + $u2
if ($sum -ne 5000000000) {
    Write-Host "FAIL: uint32 widening addition mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
