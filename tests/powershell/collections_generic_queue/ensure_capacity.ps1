# vybe-test: powershell/collections_generic_queue/ensure_capacity
$q = [System.Collections.Generic.Queue[int]]::new()
$cap = $q.EnsureCapacity(50)
if ($cap -lt 50) {
    Write-Host "FAIL: EnsureCapacity on Queue failed, got $cap"
    exit 1
}
Write-Host "PASS"
exit 0
