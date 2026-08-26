# vybe-test: powershell/collections_bitarray_operations/length_resize_truncates
$ba = [System.Collections.BitArray]::new(@($true, $true, $true, $true))
$ba.Length = 2
if ($ba.Length -ne 2 -or $ba[0] -ne $true -or $ba[1] -ne $true) {
    Write-Host "FAIL: Length truncation failed"
    exit 1
}
Write-Host "PASS"
exit 0
