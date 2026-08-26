# vybe-test: powershell/string_encoding_hex_conversions/bitconverter_is_little_endian
$isLe = [System.BitConverter]::IsLittleEndian
if ($isLe -ne $true -and $isLe -ne $false) {
    Write-Host "FAIL: IsLittleEndian must return boolean"
    exit 1
}
Write-Host "PASS"
exit 0
