# vybe-test: powershell/collections_bitarray_operations/copyto_int_array
$ba = [System.Collections.BitArray]::new(32, $false)
$ba.Set(0, $true); $ba.Set(1, $true) # 3
[int[]]$target = New-Object int[] 1
$ba.CopyTo($target, 0)
if ($target[0] -ne 3) {
    Write-Host "FAIL: CopyTo int array failed, expected 3, got $($target[0])"
    exit 1
}
Write-Host "PASS"
exit 0
