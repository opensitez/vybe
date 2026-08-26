# vybe-test: powershell/type_unsigned_integers/sbyte_min_and_max_values
$min = [sbyte]::MinValue
$max = [sbyte]::MaxValue
if ($min -ne -128 -or $max -ne 127) {
    Write-Host "FAIL: sbyte min/max mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
