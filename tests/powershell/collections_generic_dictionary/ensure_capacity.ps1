# vybe-test: powershell/collections_generic_dictionary/ensure_capacity
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$cap = $d.EnsureCapacity(100)
if ($cap -lt 100) {
    Write-Host "FAIL: EnsureCapacity failed, got $cap"
    exit 1
}
Write-Host "PASS"
exit 0
