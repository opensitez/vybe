# vybe-test: powershell/new_object_cmdlet/new_object_queue_enqueue_dequeue_fifo
# New-Object System.Collections.Queue provides FIFO enqueue/dequeue semantics
$q = New-Object System.Collections.Queue
$q.Enqueue("first")
$q.Enqueue("second")

$dequeued = $q.Dequeue()

if ($dequeued -ne "first") {
    Write-Host "FAIL: FIFO order violation, expected 'first', got '$dequeued'"
    exit 1
}

Write-Host "PASS"
exit 0
