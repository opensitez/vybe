# vybe-test: powershell/collections_bitarray_operations/all_zeros_check
$ba = [System.Collections.BitArray]::new(16, $false)
$anyTrue = $false
foreach ($bit in $ba) {
    if ($bit) { $anyTrue = $true; break }
}
if ($anyTrue) {
    Write-Host "FAIL: All-zeros BitArray contained a true bit"
    exit 1
}
Write-Host "PASS"
exit 0
