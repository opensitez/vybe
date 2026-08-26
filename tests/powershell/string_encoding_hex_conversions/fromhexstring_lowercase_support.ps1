# vybe-test: powershell/string_encoding_hex_conversions/fromhexstring_lowercase_support
$hex = "deadbeef"
$bytes = [System.Convert]::FromHexString($hex)
if ($bytes[0] -ne 0xDE -or $bytes[1] -ne 0xAD -or $bytes[2] -ne 0xBE -or $bytes[3] -ne 0xEF) {
    Write-Host "FAIL: Lowercase FromHexString failed"
    exit 1
}
Write-Host "PASS"
exit 0
