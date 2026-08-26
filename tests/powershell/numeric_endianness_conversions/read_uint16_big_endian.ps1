# vybe-test: powershell/numeric_endianness_conversions/read_uint16_big_endian
[byte[]]$bytes = @(0x01, 0x02)
$val = [System.Buffers.Binary.BinaryPrimitives]::ReadUInt16BigEndian($bytes)
if ($val -ne 0x0102) {
    Write-Host "FAIL: ReadUInt16BigEndian expected 0x0102, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
