# vybe-test: powershell/type_unsigned_integers/uint64_min_and_max_values
$min = [uint64]::MinValue
$max = [uint64]::MaxValue
if ($min -ne 0 -or $max -ne 18446744073709551615) {
    Write-Host "FAIL: uint64 min/max mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
