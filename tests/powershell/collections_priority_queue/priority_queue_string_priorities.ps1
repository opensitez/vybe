# vybe-test: powershell/collections_priority_queue/priority_queue_string_priorities
$pq = [System.Collections.Generic.PriorityQueue[int, string]]::new()
$pq.Enqueue(100, "b")
$pq.Enqueue(200, "a")
if ($pq.Dequeue() -ne 200) { Write-Host "FAIL: String priority ordering failed"; exit 1 }
Write-Host "PASS"; exit 0
