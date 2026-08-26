# vybe-test: powershell/numeric_endianness_conversions/read_int32_big_endian_from_bytes
[byte[]]$beBytes = @(0x12, 0x34, 0x56, 0x78)
$val = [System.Buffers.Binary.BinaryPrimitives]::ReadInt32BigEndian($beBytes)
if ($val -ne 0x12345678) {
    Write-Host "FAIL: ReadInt32BigEndian failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
