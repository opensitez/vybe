# vybe-test: powershell/string_encoding_hex_conversions/int32_to_hex_formatting
$val = 4096
$hex = $val.ToString("X4")
if ($hex -ne "1000") {
    Write-Host "FAIL: Int32 to hex format failed, expected '1000', got '$hex'"
    exit 1
}
Write-Host "PASS"
exit 0
