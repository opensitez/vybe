# vybe-test: powershell/numeric_endianness_conversions/reverse_endianness_uint32
$val = [uint32]0x12345678
$bytes = [System.BitConverter]::GetBytes($val)
[System.Array]::Reverse($bytes)
$reversed = [System.BitConverter]::ToUInt32($bytes, 0)
if ($reversed -ne [uint32]0x78563412) {
    Write-Host "FAIL: ReverseEndianness uint32 failed"
    exit 1
}
Write-Host "PASS"
exit 0
