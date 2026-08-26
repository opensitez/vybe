# vybe-test: powershell/collections_generic_stack/invalidoperationexception_on_empty_pop
$s = [System.Collections.Generic.Stack[int]]::new()
$caught = $false
try {
    $x = $s.Pop()
} catch [System.InvalidOperationException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Pop on empty stack must throw InvalidOperationException"
    exit 1
}
Write-Host "PASS"
exit 0
