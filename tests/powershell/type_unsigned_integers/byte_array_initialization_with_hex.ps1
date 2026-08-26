# vybe-test: powershell/type_unsigned_integers/byte_array_initialization_with_hex
[byte[]]$bytes = @(0x10, 0x20, 0x30, 0x40)
if ($bytes.Length -ne 4 -or $bytes[0] -ne 16 -or $bytes[3] -ne 64) {
    Write-Host "FAIL: byte array initialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
