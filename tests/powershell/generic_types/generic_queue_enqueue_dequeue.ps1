# vybe-test: powershell/generic_types/generic_queue_enqueue_dequeue
$q = [System.Collections.Generic.Queue[string]]::new()
$q.Enqueue("First")
$q.Enqueue("Second")
$item = $q.Dequeue()
if ($item -ne "First" -or $q.Count -ne 1) {
    Write-Host "FAIL: Queue Dequeue expected First, remaining Count 1"
    exit 1
}
Write-Host "PASS"
exit 0
