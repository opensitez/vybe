# vybe-test: powershell/string_encoding_hex_conversions/bitconverter_toint64_from_hex
$hex = "0100000000000000"
$bytes = [System.Convert]::FromHexString($hex)
$val = [System.BitConverter]::ToInt64($bytes, 0)
if ($val -ne 1) {
    Write-Host "FAIL: BitConverter ToInt64 from hex failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
