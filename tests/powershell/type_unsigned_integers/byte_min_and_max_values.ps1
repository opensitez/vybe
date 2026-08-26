# vybe-test: powershell/type_unsigned_integers/byte_min_and_max_values
$min = [byte]::MinValue
$max = [byte]::MaxValue
if ($min -ne 0 -or $max -ne 255) {
    Write-Host "FAIL: byte min/max mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
