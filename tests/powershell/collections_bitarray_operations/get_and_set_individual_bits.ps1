# vybe-test: powershell/collections_bitarray_operations/get_and_set_individual_bits
$ba = [System.Collections.BitArray]::new(4)
$ba.Set(1, $true)
$ba.Set(3, $true)
if ($ba[0] -ne $false -or $ba[1] -ne $true -or $ba[2] -ne $false -or $ba[3] -ne $true) {
    Write-Host "FAIL: BitArray Get/Set failed"
    exit 1
}
Write-Host "PASS"
exit 0
