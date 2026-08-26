# vybe-test: powershell/collections_bitarray_operations/copyto_byte_array
$ba = [System.Collections.BitArray]::new(8, $false)
$ba.Set(0, $true); $ba.Set(2, $true) # 0b00000101 = 5
[byte[]]$target = New-Object byte[] 1
$ba.CopyTo($target, 0)
if ($target[0] -ne 5) {
    Write-Host "FAIL: CopyTo byte array failed, expected 5, got $($target[0])"
    exit 1
}
Write-Host "PASS"
exit 0
