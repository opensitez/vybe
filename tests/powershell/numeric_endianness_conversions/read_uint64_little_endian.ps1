# vybe-test: powershell/numeric_endianness_conversions/read_uint64_little_endian
[byte[]]$bytes = @(1, 0, 0, 0, 0, 0, 0, 0)
$val = [System.Buffers.Binary.BinaryPrimitives]::ReadUInt64LittleEndian($bytes)
if ($val -ne 1) {
    Write-Host "FAIL: ReadUInt64LittleEndian expected 1, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
