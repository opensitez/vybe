# vybe-test: powershell/collections_generic_stack/ensure_capacity
$s = [System.Collections.Generic.Stack[int]]::new()
$cap = $s.EnsureCapacity(60)
if ($cap -lt 60) {
    Write-Host "FAIL: EnsureCapacity on Stack failed, got $cap"
    exit 1
}
Write-Host "PASS"
exit 0
