# vybe-test: powershell/collections_priority_queue/priority_queue_ensure_capacity
$pq = [System.Collections.Generic.PriorityQueue[int, int]]::new()
$cap = $pq.EnsureCapacity(100)
if ($cap -lt 100) { Write-Host "FAIL: EnsureCapacity failed"; exit 1 }
Write-Host "PASS"; exit 0
