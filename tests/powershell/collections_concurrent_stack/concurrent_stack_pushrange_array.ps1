# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_pushrange_array
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new()
$cs.PushRange([int[]]@(10, 20, 30))
[int]$top = 0
$null = $cs.TryPop([ref]$top)
if ($top -ne 30 -or $cs.Count -ne 2) { Write-Host "FAIL: PushRange failed"; exit 1 }
Write-Host "PASS"; exit 0
