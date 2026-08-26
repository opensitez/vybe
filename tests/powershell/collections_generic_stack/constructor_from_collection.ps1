# vybe-test: powershell/collections_generic_stack/constructor_from_collection
[int[]]$init = @(1, 2, 3)
$s = [System.Collections.Generic.Stack[int]]::new($init)
$popped = $s.Pop()
if ($popped -ne 3 -or $s.Count -ne 2) {
    Write-Host "FAIL: Stack constructor from collection failed"
    exit 1
}
Write-Host "PASS"
exit 0
