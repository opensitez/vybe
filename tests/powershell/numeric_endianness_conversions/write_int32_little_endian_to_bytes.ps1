# vybe-test: powershell/numeric_endianness_conversions/write_int32_little_endian_to_bytes
[byte[]]$buf = New-Object byte[] 4
[System.Buffers.Binary.BinaryPrimitives]::WriteInt32LittleEndian($buf, 0x12345678)
if ($buf[0] -ne 0x78 -or $buf[1] -ne 0x56 -or $buf[2] -ne 0x34 -or $buf[3] -ne 0x12) {
    Write-Host "FAIL: WriteInt32LittleEndian failed"
    exit 1
}
Write-Host "PASS"
exit 0
