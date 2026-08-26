# vybe-test: powershell/collections_bitarray_operations/and_operation
$ba1 = [System.Collections.BitArray]::new(@($true, $true, $false, $false))
$ba2 = [System.Collections.BitArray]::new(@($true, $false, $true, $false))
$ba1.And($ba2)
if ($ba1[0] -ne $true -or $ba1[1] -ne $false -or $ba1[2] -ne $false -or $ba1[3] -ne $false) {
    Write-Host "FAIL: BitArray And failed"
    exit 1
}
Write-Host "PASS"
exit 0
