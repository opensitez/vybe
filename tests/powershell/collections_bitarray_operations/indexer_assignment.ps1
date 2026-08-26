# vybe-test: powershell/collections_bitarray_operations/indexer_assignment
$ba = [System.Collections.BitArray]::new(2, $false)
$ba[1] = $true
if ($ba[0] -ne $false -or $ba[1] -ne $true) {
    Write-Host "FAIL: Indexer assignment failed"
    exit 1
}
Write-Host "PASS"
exit 0
