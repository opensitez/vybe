# vybe-test: powershell/collections/queue_enqueue_dequeue
$q = [System.Collections.Generic.Queue[int]]::new()
$q.Enqueue(1)
$q.Enqueue(2)
$q.Enqueue(3)
$first = $q.Dequeue()
if ($first -ne 1) { Write-Host "FAIL: FIFO dequeue, expected 1 got $first"; exit 1 }
if ($q.Count -ne 2) { Write-Host "FAIL: count"; exit 1 }
if ($q.Peek() -ne 2) { Write-Host "FAIL: peek after dequeue"; exit 1 }
Write-Host "PASS"
exit 0
