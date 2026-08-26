# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_push_and_trypop
$cs = [System.Collections.Concurrent.ConcurrentStack[int]]::new()
$cs.Push(10)
[int]$outVal = 0
$ok = $cs.TryPop([ref]$outVal)
if (-not $ok -or $outVal -ne 10 -or $cs.Count -ne 0) { Write-Host "FAIL: ConcurrentStack Push/TryPop failed"; exit 1 }
Write-Host "PASS"; exit 0
