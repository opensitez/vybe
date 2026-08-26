# vybe-test: powershell/math_ieee_special_values/bit_converter_double_to_int64_bits
$bits = [System.BitConverter]::DoubleToInt64Bits(0.0)
$negBits = [System.BitConverter]::DoubleToInt64Bits(-0.0)
if ($bits -ne 0 -or $negBits -eq 0) {
    Write-Host "FAIL: DoubleToInt64Bits distinguishing -0.0 failed"
    exit 1
}
Write-Host "PASS"
exit 0
