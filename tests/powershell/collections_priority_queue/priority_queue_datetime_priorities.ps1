# vybe-test: powershell/collections_priority_queue/priority_queue_datetime_priorities
$pq = [System.Collections.Generic.PriorityQueue[string, datetime]]::new()
$now = [datetime]::UtcNow
$pq.Enqueue("Later", $now.AddHours(2))
$pq.Enqueue("Earlier", $now.AddHours(1))
if ($pq.Dequeue() -ne "Earlier") { Write-Host "FAIL: DateTime priority ordering failed"; exit 1 }
Write-Host "PASS"; exit 0
