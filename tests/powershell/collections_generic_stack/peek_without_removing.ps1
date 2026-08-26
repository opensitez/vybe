# vybe-test: powershell/collections_generic_stack/peek_without_removing
$s = [System.Collections.Generic.Stack[string]]::new()
$s.Push("item")
$peeked = $s.Peek()
if ($peeked -ne "item" -or $s.Count -ne 1) {
    Write-Host "FAIL: Stack Peek mutated count"
    exit 1
}
Write-Host "PASS"
exit 0
