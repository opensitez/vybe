# vybe-test: powershell/string_encoding_hex_conversions/fromhexstring_empty_string
$bytes = [System.Convert]::FromHexString("")
if ($bytes.Length -ne 0) {
    Write-Host "FAIL: FromHexString empty string failed"
    exit 1
}
Write-Host "PASS"
exit 0
