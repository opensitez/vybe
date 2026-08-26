# vybe-test: powershell/type_unsigned_integers/uint64_large_value_arithmetic
[uint64]$u = 10000000000000000000
$half = $u / 2
if ($half -ne 5000000000000000000) {
    Write-Host "FAIL: uint64 division mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
