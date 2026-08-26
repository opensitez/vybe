# vybe-test: powershell/collections_priority_queue/priority_queue_with_duplicate_priorities_fifo_or_valid
$pq = [System.Collections.Generic.PriorityQueue[string, int]]::new()
$pq.Enqueue("T1", 5); $pq.Enqueue("T2", 5)
$d1 = $pq.Dequeue()
$d2 = $pq.Dequeue()
if ($pq.Count -ne 0 -or ($d1 -ne "T1" -and $d1 -ne "T2")) { Write-Host "FAIL: Duplicate priority dequeue failed"; exit 1 }
Write-Host "PASS"; exit 0
