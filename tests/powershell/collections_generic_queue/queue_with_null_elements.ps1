# vybe-test: powershell/collections_generic_queue/queue_with_null_elements
$q = [System.Collections.Generic.Queue[string]]::new()
$q.Enqueue("hello")
$item = $q.Dequeue()
if ($item -ne "hello" -or $q.Count -ne 0) {
    Write-Host "FAIL: Queue Dequeue failed"
    exit 1
}
Write-Host "PASS"
exit 0
