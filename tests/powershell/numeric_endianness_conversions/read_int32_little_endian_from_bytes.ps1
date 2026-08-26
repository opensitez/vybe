# vybe-test: powershell/numeric_endianness_conversions/read_int32_little_endian_from_bytes
[byte[]]$leBytes = @(0x78, 0x56, 0x34, 0x12)
$val = [System.Buffers.Binary.BinaryPrimitives]::ReadInt32LittleEndian($leBytes)
if ($val -ne 0x12345678) {
    Write-Host "FAIL: ReadInt32LittleEndian failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
