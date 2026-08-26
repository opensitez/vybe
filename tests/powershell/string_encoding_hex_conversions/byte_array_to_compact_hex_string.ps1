# vybe-test: powershell/string_encoding_hex_conversions/byte_array_to_compact_hex_string
[byte[]]$bytes = @(10, 20, 30)
$hex = -join ($bytes | ForEach-Object { $_.ToString("X2") })
if ($hex -ne "0A141E") {
    Write-Host "FAIL: Compact hex string creation failed, got '$hex'"
    exit 1
}
Write-Host "PASS"
exit 0
