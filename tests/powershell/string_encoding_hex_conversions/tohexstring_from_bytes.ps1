# vybe-test: powershell/string_encoding_hex_conversions/tohexstring_from_bytes
[byte[]]$bytes = @(0xDE, 0xAD, 0xBE, 0xEF)
$hex = [System.Convert]::ToHexString($bytes)
if ($hex -ne "DEADBEEF") {
    Write-Host "FAIL: ToHexString failed, expected 'DEADBEEF', got '$hex'"
    exit 1
}
Write-Host "PASS"
exit 0
