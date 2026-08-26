# vybe-test: powershell/collections_bitarray_operations/is_synchronized_and_syncroot
$ba = [System.Collections.BitArray]::new(4)
if ($ba.IsSynchronized -ne $false -or $ba.SyncRoot -eq $null) {
    Write-Host "FAIL: SyncRoot check failed"
    exit 1
}
Write-Host "PASS"
exit 0
