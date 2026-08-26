# vybe-test: powershell/collections_priority_queue/priority_queue_enqueue_dequeue_fused_operation
$pq = [System.Collections.Generic.PriorityQueue[string, int]]::new()
$pq.Enqueue("OldMin", 5)
$res = $pq.EnqueueDequeue("NewItem", 10)
if ($res -ne "OldMin" -or $pq.Peek() -ne "NewItem") { Write-Host "FAIL: EnqueueDequeue failed"; exit 1 }
Write-Host "PASS"; exit 0
