# vybe-test: powershell/collections_generic_queue/drain_all_with_while_loop
$q = [System.Collections.Generic.Queue[int]]::new()
for ($i = 0; $i -lt 5; $i++) { $q.Enqueue($i) }
$drained = @()
while ($q.Count -gt 0) { $drained += $q.Dequeue() }
if ($drained.Count -ne 5 -or $drained[0] -ne 0 -or $drained[4] -ne 4 -or $q.Count -ne 0) {
    Write-Host "FAIL: Drain queue loop failed"
    exit 1
}
Write-Host "PASS"
exit 0
