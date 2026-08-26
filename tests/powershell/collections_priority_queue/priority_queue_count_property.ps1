# vybe-test: powershell/collections_priority_queue/priority_queue_count_property
$pq = [System.Collections.Generic.PriorityQueue[int, int]]::new()
for ($i = 0; $i -lt 15; $i++) { $pq.Enqueue($i, $i) }
if ($pq.Count -ne 15) { Write-Host "FAIL: Count property failed"; exit 1 }
Write-Host "PASS"; exit 0
