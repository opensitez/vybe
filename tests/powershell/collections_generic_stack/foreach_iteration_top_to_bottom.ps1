# vybe-test: powershell/collections_generic_stack/foreach_iteration_top_to_bottom
$s = [System.Collections.Generic.Stack[int]]::new()
$s.Push(1); $s.Push(2); $s.Push(3)
$items = @()
foreach ($item in $s) { $items += $item }
if ($items[0] -ne 3 -or $items[2] -ne 1 -or $s.Count -ne 3) {
    Write-Host "FAIL: Foreach on stack order failed"
    exit 1
}
Write-Host "PASS"
exit 0
