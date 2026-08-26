# vybe-test: powershell/numeric_endianness_conversions/write_double_big_endian
[byte[]]$buf = New-Object byte[] 8
[System.Buffers.Binary.BinaryPrimitives]::WriteDoubleBigEndian($buf, 1.0)
if ($buf[0] -ne 0x3F -or $buf[1] -ne 0xF0) {
    Write-Host "FAIL: WriteDoubleBigEndian failed"
    exit 1
}
Write-Host "PASS"
exit 0
