# vybe-test: powershell/collections_bitarray_operations/construct_from_bytes
[byte[]]$bytes = @(0x05) # binary: 00000101 (LSB is bit 0: true, bit 1: false, bit 2: true)
$ba = [System.Collections.BitArray]::new($bytes)
if ($ba.Length -ne 8 -or $ba[0] -ne $true -or $ba[1] -ne $false -or $ba[2] -ne $true -or $ba[3] -ne $false) {
    Write-Host "FAIL: BitArray from bytes failed"
    exit 1
}
Write-Host "PASS"
exit 0
