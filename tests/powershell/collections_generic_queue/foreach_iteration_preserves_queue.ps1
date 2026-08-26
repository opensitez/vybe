# vybe-test: powershell/collections_generic_queue/foreach_iteration_preserves_queue
$q = [System.Collections.Generic.Queue[int]]::new()
$q.Enqueue(10); $q.Enqueue(20); $q.Enqueue(30)
$sum = 0
foreach ($item in $q) { $sum += $item }
if ($sum -ne 60 -or $q.Count -ne 3) {
    Write-Host "FAIL: Foreach on queue mutated count or computed incorrect sum"
    exit 1
}
Write-Host "PASS"
exit 0
