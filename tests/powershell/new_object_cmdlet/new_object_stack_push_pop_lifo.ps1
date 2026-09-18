# vybe-test: powershell/new_object_cmdlet/new_object_stack_push_pop_lifo
# New-Object System.Collections.Stack provides LIFO push/pop semantics
$stack = New-Object System.Collections.Stack
$stack.Push("first")
$stack.Push("second")

$popped = $stack.Pop()

if ($popped -ne "second") {
    Write-Host "FAIL: LIFO order violation, expected 'second', got '$popped'"
    exit 1
}

if ($stack.Count -ne 1) {
    Write-Host "FAIL: expected Count 1 after pop, got $($stack.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
