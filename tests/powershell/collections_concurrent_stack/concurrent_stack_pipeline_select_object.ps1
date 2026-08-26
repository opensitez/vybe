# vybe-test: powershell/collections_concurrent_stack/concurrent_stack_pipeline_select_object
$cs = [System.Collections.Concurrent.ConcurrentStack[string]]::new([string[]]@("a", "b"))
$res = @($cs | Select-Object @{ N = "Upper"; E = { $_.ToUpper() } })
if ($res.Length -ne 2) { Write-Host "FAIL: Pipeline Select-Object failed"; exit 1 }
Write-Host "PASS"; exit 0
