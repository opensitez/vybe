# vybe-test: powershell/collections_generic_dictionary/tryadd_method_powershell
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$ok1 = $d.TryAdd("k", 1)
$ok2 = $d.TryAdd("k", 2)
if (-not $ok1 -or $ok2 -or $d["k"] -ne 1) {
    Write-Host "FAIL: TryAdd method failed"
    exit 1
}
Write-Host "PASS"
exit 0
