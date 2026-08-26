# vybe-test: powershell/string_encoding_hex_conversions/bitconverter_getbytes_int32
$val = 258 # 0x00000102
$bytes = [System.BitConverter]::GetBytes([int]$val)
if ($bytes[0] -ne 0x02 -or $bytes[1] -ne 0x01) {
    Write-Host "FAIL: BitConverter GetBytes failed"
    exit 1
}
Write-Host "PASS"
exit 0
