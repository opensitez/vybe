# vybe-test: powershell/collections_bitarray_operations/foreach_enumeration
$ba = [System.Collections.BitArray]::new(@($true, $false, $true))
$trueCount = 0
foreach ($b in $ba) {
    if ($b -eq $true) { $trueCount++ }
}
if ($trueCount -ne 2) {
    Write-Host "FAIL: Foreach true count mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
