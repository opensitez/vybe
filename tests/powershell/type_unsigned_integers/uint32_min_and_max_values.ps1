# vybe-test: powershell/type_unsigned_integers/uint32_min_and_max_values
$min = [uint32]::MinValue
$max = [uint32]::MaxValue
if ($min -ne 0 -or $max -ne 4294967295) {
    Write-Host "FAIL: uint32 min/max mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
