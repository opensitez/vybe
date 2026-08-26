# vybe-test: powershell/collections_bitarray_operations/construct_with_length_and_default_value
$ba = [System.Collections.BitArray]::new(8, $false)
if ($ba.Length -ne 8 -or $ba[0] -ne $false -or $ba[7] -ne $false) {
    Write-Host "FAIL: BitArray initial length/value failed"
    exit 1
}
Write-Host "PASS"
exit 0
