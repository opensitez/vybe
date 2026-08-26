# vybe-test: powershell/collections_priority_queue/priority_queue_clear_all
$pq = [System.Collections.Generic.PriorityQueue[string, int]]::new()
$pq.Enqueue("A", 1); $pq.Enqueue("B", 2)
$pq.Clear()
if ($pq.Count -ne 0) { Write-Host "FAIL: PriorityQueue Clear failed"; exit 1 }
Write-Host "PASS"; exit 0
