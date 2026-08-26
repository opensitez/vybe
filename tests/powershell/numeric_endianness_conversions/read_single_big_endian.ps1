# vybe-test: powershell/numeric_endianness_conversions/read_single_big_endian
[byte[]]$bytes = @(0x3F, 0x80, 0x00, 0x00) # 1.0f in IEEE 754 Big Endian
$val = [System.Buffers.Binary.BinaryPrimitives]::ReadSingleBigEndian($bytes)
if ($val -ne 1.0) {
    Write-Host "FAIL: ReadSingleBigEndian expected 1.0, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
