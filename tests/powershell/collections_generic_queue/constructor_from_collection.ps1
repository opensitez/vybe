# vybe-test: powershell/collections_generic_queue/constructor_from_collection
[int[]]$init = @(5, 6, 7)
$q = [System.Collections.Generic.Queue[int]]::new($init)
if ($q.Count -ne 3 -or $q.Dequeue() -ne 5) {
    Write-Host "FAIL: Queue initialization from collection failed"
    exit 1
}
Write-Host "PASS"
exit 0
