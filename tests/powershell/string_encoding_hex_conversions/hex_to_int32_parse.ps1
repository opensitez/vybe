# vybe-test: powershell/string_encoding_hex_conversions/hex_to_int32_parse
$hex = "000000FF"
$val = [int]::Parse($hex, [System.Globalization.NumberStyles]::HexNumber)
if ($val -ne 255) {
    Write-Host "FAIL: Hex to int32 parse failed"
    exit 1
}
Write-Host "PASS"
exit 0
