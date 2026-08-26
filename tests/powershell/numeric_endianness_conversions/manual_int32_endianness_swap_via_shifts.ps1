# vybe-test: powershell/numeric_endianness_conversions/manual_int32_endianness_swap_via_shifts
[int32]$val = 0x12345678
$b0 = ($val -shr 0) -band 0xFF
$b1 = ($val -shr 8) -band 0xFF
$b2 = ($val -shr 16) -band 0xFF
$b3 = ($val -shr 24) -band 0xFF
$swapped = ($b0 -shl 24) -bor ($b1 -shl 16) -bor ($b2 -shl 8) -bor $b3
if ($swapped -ne 0x78563412) {
    Write-Host "FAIL: Manual shift endianness swap failed, got $swapped"
    exit 1
}
Write-Host "PASS"
exit 0
