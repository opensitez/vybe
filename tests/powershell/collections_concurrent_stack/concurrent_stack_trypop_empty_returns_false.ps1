# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_trypop_empty_returns_false
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new()
[int]$outVal = 0
$ok = $cs.TryPop([ref]$outVal)
if ($ok) { Write-Host "FAIL: TryPop empty should return false"; exit 1 }
Write-Host "PASS"; exit 0
