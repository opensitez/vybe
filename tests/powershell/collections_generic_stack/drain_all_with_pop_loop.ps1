# vybe-test: powershell/collections_generic_stack/drain_all_with_pop_loop
$s = [System.Collections.Generic.Stack[int]]::new()
for ($i = 1; $i -le 5; $i++) { $s.Push($i) }
$popped = @()
while ($s.Count -gt 0) { $popped += $s.Pop() }
if ($popped[0] -ne 5 -or $popped[4] -ne 1 -or $s.Count -ne 0) {
    Write-Host "FAIL: Drain stack loop failed"
    exit 1
}
Write-Host "PASS"
exit 0
