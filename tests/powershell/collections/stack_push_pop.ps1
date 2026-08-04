# vybe-test: powershell/collections/stack_push_pop
$stack = [System.Collections.Generic.Stack[string]]::new()
$stack.Push("first")
$stack.Push("second")
$stack.Push("third")
$top = $stack.Peek()
if ($top -ne "third") { Write-Host "FAIL: peek"; exit 1 }
$popped = $stack.Pop()
if ($popped -ne "third") { Write-Host "FAIL: pop"; exit 1 }
if ($stack.Count -ne 2) { Write-Host "FAIL: count after pop"; exit 1 }
Write-Host "PASS"
exit 0
