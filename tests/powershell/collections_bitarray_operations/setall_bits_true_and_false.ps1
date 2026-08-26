# vybe-test: powershell/collections_bitarray_operations/setall_bits_true_and_false
$ba = [System.Collections.BitArray]::new(5)
$ba.SetAll($true)
if ($ba[0] -ne $true -or $ba[4] -ne $true) {
    Write-Host "FAIL: SetAll($true) failed"
    exit 1
}
$ba.SetAll($false)
if ($ba[0] -ne $false -or $ba[4] -ne $false) {
    Write-Host "FAIL: SetAll($false) failed"
    exit 1
}
Write-Host "PASS"
exit 0
