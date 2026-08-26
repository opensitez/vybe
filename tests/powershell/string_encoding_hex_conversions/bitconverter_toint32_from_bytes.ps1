# vybe-test: powershell/string_encoding_hex_conversions/bitconverter_toint32_from_bytes
[byte[]]$bytes = @(0x2A, 0x00, 0x00, 0x00) # 42 in little-endian
$val = [System.BitConverter]::ToInt32($bytes, 0)
if ($val -ne 42) {
    Write-Host "FAIL: BitConverter ToInt32 failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
