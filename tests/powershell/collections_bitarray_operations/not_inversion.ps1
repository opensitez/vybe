# vybe-test: powershell/collections_bitarray_operations/not_inversion
$ba = [System.Collections.BitArray]::new(@($true, $false, $true))
$ba.Not()
if ($ba[0] -ne $false -or $ba[1] -ne $true -or $ba[2] -ne $false) {
    Write-Host "FAIL: BitArray Not failed"
    exit 1
}
Write-Host "PASS"
exit 0
