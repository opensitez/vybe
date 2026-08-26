# vybe-test: powershell/collections_generic_stack/clear_stack
$s = [System.Collections.Generic.Stack[int]]::new()
$s.Push(1); $s.Push(2)
$s.Clear()
if ($s.Count -ne 0) {
    Write-Host "FAIL: Stack Clear failed"
    exit 1
}
Write-Host "PASS"
exit 0
