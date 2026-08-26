# vybe-test: powershell/numeric_endianness_conversions/reverse_endianness_int16
[int16]$val = 0x1234
[int16]$rev = [System.Buffers.Binary.BinaryPrimitives]::ReverseEndianness($val)
if ($rev -ne 0x3412) {
    Write-Host "FAIL: ReverseEndianness int16 failed, expected 0x3412, got $rev"
    exit 1
}
Write-Host "PASS"
exit 0
