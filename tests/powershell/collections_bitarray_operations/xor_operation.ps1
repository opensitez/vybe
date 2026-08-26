# vybe-test: powershell/collections_bitarray_operations/xor_operation
$ba1 = [System.Collections.BitArray]::new(@($true, $true, $false))
$ba2 = [System.Collections.BitArray]::new(@($true, $false, $false))
$ba1.Xor($ba2)
if ($ba1[0] -ne $false -or $ba1[1] -ne $true -or $ba1[2] -ne $false) {
    Write-Host "FAIL: BitArray Xor failed"
    exit 1
}
Write-Host "PASS"
exit 0
