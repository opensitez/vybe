# vybe-test: powershell/numeric_endianness_conversions/reverse_endianness_int64
[int64]$val = 0x0102030405060708
[int64]$rev = [System.Buffers.Binary.BinaryPrimitives]::ReverseEndianness($val)
if ($rev -ne 0x0807060504030201) {
    Write-Host "FAIL: ReverseEndianness int64 failed"
    exit 1
}
Write-Host "PASS"
exit 0
