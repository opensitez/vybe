# vybe-test: powershell/numeric_endianness_conversions/reverse_endianness_int32
[int32]$val = 0x12345678
[int32]$rev = [System.Buffers.Binary.BinaryPrimitives]::ReverseEndianness($val)
if ($rev -ne 0x78563412) {
    Write-Host "FAIL: ReverseEndianness int32 failed, expected 0x78563412, got $rev"
    exit 1
}
Write-Host "PASS"
exit 0
