# vybe-test: powershell/type_biginteger_arithmetic/parse_from_hex_string
$hexStr = "00FF"
$val = [bigint]::Parse($hexStr, [System.Globalization.NumberStyles]::HexNumber)
if ($val -ne [bigint]255) {
    Write-Host "FAIL: expected 255, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
