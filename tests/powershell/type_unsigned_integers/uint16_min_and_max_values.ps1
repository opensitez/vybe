# vybe-test: powershell/type_unsigned_integers/uint16_min_and_max_values
$min = [uint16]::MinValue
$max = [uint16]::MaxValue
if ($min -ne 0 -or $max -ne 65535) {
    Write-Host "FAIL: uint16 min/max mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
