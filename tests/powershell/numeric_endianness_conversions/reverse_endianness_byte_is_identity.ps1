# vybe-test: powershell/numeric_endianness_conversions/reverse_endianness_byte_is_identity
[byte]$b = 0x42
[byte]$rev = [System.Buffers.Binary.BinaryPrimitives]::ReverseEndianness($b)
if ($b -ne $rev) {
    Write-Host "FAIL: Byte reverse endianness should be identity"
    exit 1
}
Write-Host "PASS"
exit 0
