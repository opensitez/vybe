# vybe-test: powershell/collections_bitarray_operations/clone_bitarray
$b1 = [System.Collections.BitArray]::new(@($true, $false))
$b2 = $b1.Clone()
$b1.Set(0, $false)
if ($b2[0] -ne $true) {
    Write-Host "FAIL: Clone should be independent copy"
    exit 1
}
Write-Host "PASS"
exit 0
