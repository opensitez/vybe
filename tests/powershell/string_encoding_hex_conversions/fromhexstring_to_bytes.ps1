# vybe-test: powershell/string_encoding_hex_conversions/fromhexstring_to_bytes
$hex = "CAFEBABE"
$bytes = [System.Convert]::FromHexString($hex)
if ($bytes.Length -ne 4 -or $bytes[0] -ne 0xCA -or $bytes[3] -ne 0xBE) {
    Write-Host "FAIL: FromHexString failed"
    exit 1
}
Write-Host "PASS"
exit 0
