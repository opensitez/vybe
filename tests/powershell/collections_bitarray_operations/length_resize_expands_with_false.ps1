# vybe-test: powershell/collections_bitarray_operations/length_resize_expands_with_false
$ba = [System.Collections.BitArray]::new(@($true, $true))
$ba.Length = 4
if ($ba.Length -ne 4 -or $ba[0] -ne $true -or $ba[1] -ne $true -or $ba[2] -ne $false -or $ba[3] -ne $false) {
    Write-Host "FAIL: Length expansion failed"
    exit 1
}
Write-Host "PASS"
exit 0
