# vybe-test: powershell/bitwise_rotation_and_shifts/high_and_low_nibble_extraction
[byte]$b = 0xAB
$high = ($b -shr 4) -band 0x0F
$low = $b -band 0x0F
if ($high -ne 0x0A -or $low -ne 0x0B) {
    Write-Host "FAIL: High/low nibble extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
