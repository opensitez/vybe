# vybe-test: powershell/collections_priority_queue/priority_queue_trim_excess
$pq = [System.Collections.Generic.PriorityQueue[int, int]]::new(100)
$pq.Enqueue(1, 1)
$pq.TrimExcess()
if ($pq.Count -ne 1) { Write-Host "FAIL: TrimExcess failed"; exit 1 }
Write-Host "PASS"; exit 0
