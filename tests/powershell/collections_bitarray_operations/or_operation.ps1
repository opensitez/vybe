# vybe-test: powershell/collections_bitarray_operations/or_operation
$ba1 = [System.Collections.BitArray]::new(@($true, $false, $false, $false))
$ba2 = [System.Collections.BitArray]::new(@($false, $true, $false, $false))
$ba1.Or($ba2)
if ($ba1[0] -ne $true -or $ba1[1] -ne $true -or $ba1[2] -ne $false) {
    Write-Host "FAIL: BitArray Or failed"
    exit 1
}
Write-Host "PASS"
exit 0
