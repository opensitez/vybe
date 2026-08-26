# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_lifo_ordering
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new()
$cs.Push(1); $cs.Push(2); $cs.Push(3)
[int]$v1 = 0; [int]$v2 = 0; [int]$v3 = 0
$null = $cs.TryPop([ref]$v1)
$null = $cs.TryPop([ref]$v2)
$null = $cs.TryPop([ref]$v3)
if ($v1 -ne 3 -or $v2 -ne 2 -or $v3 -ne 1) { Write-Host "FAIL: LIFO ordering failed"; exit 1 }
Write-Host "PASS"; exit 0
