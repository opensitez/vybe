# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_complex_objects
$cs = [System.Collections.Concurrent.ConcurrentStack[pscustomobject]]::new()
$cs.Push([pscustomobject]@{ Score = 99 })
[pscustomobject]$outObj = $null
$ok = $cs.TryPop([ref]$outObj)
if (-not $ok -or $outObj.Score -ne 99) { Write-Host "FAIL: Complex object in ConcurrentStack failed"; exit 1 }
Write-Host "PASS"; exit 0
