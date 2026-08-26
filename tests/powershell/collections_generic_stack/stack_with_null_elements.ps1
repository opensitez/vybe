# vybe-test: powershell/collections_generic_stack/stack_with_null_elements
$s = [System.Collections.Generic.Stack[string]]::new()
$s.Push("hello")
$item = $s.Pop()
if ($item -ne "hello" -or $s.Count -ne 0) {
    Write-Host "FAIL: Stack Pop failed"
    exit 1
}
Write-Host "PASS"
exit 0
