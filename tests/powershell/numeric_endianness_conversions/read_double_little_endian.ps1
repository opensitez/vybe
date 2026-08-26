# vybe-test: powershell/numeric_endianness_conversions/read_double_little_endian
[byte[]]$bytes = @(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xF0, 0x3F) # 1.0 in IEEE 754 Little Endian
$val = [System.Buffers.Binary.BinaryPrimitives]::ReadDoubleLittleEndian($bytes)
if ($val -ne 1.0) {
    Write-Host "FAIL: ReadDoubleLittleEndian expected 1.0, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
