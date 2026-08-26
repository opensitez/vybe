# vybe-test: powershell/collections_priority_queue/priority_queue_complex_element_objects
$pq = [System.Collections.Generic.PriorityQueue[pscustomobject, int]]::new()
$pq.Enqueue([pscustomobject]@{ Task = "Build" }, 2)
$pq.Enqueue([pscustomobject]@{ Task = "Deploy" }, 1)
$res = $pq.Dequeue()
if ($res.Task -ne "Deploy") { Write-Host "FAIL: Complex object element failed"; exit 1 }
Write-Host "PASS"; exit 0
