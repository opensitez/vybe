# vybe-test: powershell/collections_generic_hashset/ensure_capacity
$set = [System.Collections.Generic.HashSet[int]]::new()
$cap = $set.EnsureCapacity(64)
if ($cap -lt 64) {
    Write-Host "FAIL: EnsureCapacity on HashSet failed, got $cap"
    exit 1
}
Write-Host "PASS"
exit 0
