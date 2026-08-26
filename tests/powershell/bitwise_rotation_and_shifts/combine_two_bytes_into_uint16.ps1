# vybe-test: powershell/bitwise_rotation_and_shifts/combine_two_bytes_into_uint16
[byte]$hi = 0x12
[byte]$lo = 0x34
[uint16]$u16 = ([uint16]$hi -shl 8) -bor [uint16]$lo
if ($u16 -ne 0x1234) {
    Write-Host "FAIL: Combine bytes into uint16 failed, got $u16"
    exit 1
}
Write-Host "PASS"
exit 0
