# vybe-test: powershell/generic_types/generic_stack_push_pop
$s = [System.Collections.Generic.Stack[int]]::new()
$s.Push(10)
$s.Push(20)
$popped = $s.Pop()
if ($popped -ne 20 -or $s.Count -ne 1) {
    Write-Host "FAIL: Stack Pop expected 20, remaining Count 1"
    exit 1
}
Write-Host "PASS"
exit 0
