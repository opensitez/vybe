# vybe-test: powershell/string_encoding_hex_conversions/tohexstring_subarray_offset_and_length
[byte[]]$bytes = @(0x01, 0x02, 0x03, 0x04, 0x05)
$hex = [System.Convert]::ToHexString($bytes, 1, 3)
if ($hex -ne "020304") {
    Write-Host "FAIL: ToHexString subarray failed, expected '020304', got '$hex'"
    exit 1
}
Write-Host "PASS"
exit 0
