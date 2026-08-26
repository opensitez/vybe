# vybe-test: powershell/collections_priority_queue/priority_queue_trydequeue_empty_returns_false
$pq = [System.Collections.Generic.PriorityQueue[string, int]]::new()
[string]$outEl = ""
[int]$outPri = 0
$ok = $pq.TryDequeue([ref]$outEl, [ref]$outPri)
if ($ok) { Write-Host "FAIL: TryDequeue on empty should return false"; exit 1 }
Write-Host "PASS"; exit 0
