# vybe-test: powershell/type_biginteger_arithmetic/conversion_to_byte_array
$val = [bigint]258 # 0x0102
$bytes = $val.ToByteArray()
if ($bytes.Length -lt 2 -or $bytes[0] -ne 2 -or $bytes[1] -ne 1) {
    Write-Host "FAIL: unexpected byte array representation"
    exit 1
}
Write-Host "PASS"
exit 0
