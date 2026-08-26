# vybe-test: powershell/collections_priority_queue/priority_queue_peek_empty_throws
$pq = [System.Collections.Generic.PriorityQueue[string, int]]::new()
$caught = $false
try {
    $x = $pq.Peek()
} catch [System.InvalidOperationException] {
    $caught = $true
}
if (-not $caught) { Write-Host "FAIL: InvalidOperationException expected on empty Peek"; exit 1 }
Write-Host "PASS"; exit 0
