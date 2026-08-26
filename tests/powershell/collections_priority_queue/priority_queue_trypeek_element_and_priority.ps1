# vybe-test: powershell/collections_priority_queue/priority_queue_trypeek_element_and_priority
$pq = [System.Collections.Generic.PriorityQueue[string, int]]::new()
$pq.Enqueue("Job", 42)
[string]$outEl = ""
[int]$outPri = 0
$ok = $pq.TryPeek([ref]$outEl, [ref]$outPri)
if (-not $ok -or $outEl -ne "Job" -or $outPri -ne 42) { Write-Host "FAIL: TryPeek failed"; exit 1 }
Write-Host "PASS"; exit 0
