# vybe-test: powershell/collections_bitarray_operations/construct_from_int_array
[int[]]$ints = @(1) # bit 0 is true, 1..31 false
$ba = [System.Collections.BitArray]::new($ints)
if ($ba.Length -ne 32 -or $ba[0] -ne $true -or $ba[1] -ne $false) {
    Write-Host "FAIL: BitArray from int array failed"
    exit 1
}
Write-Host "PASS"
exit 0
