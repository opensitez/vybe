# vybe-test: powershell/collections_generic_stack/invalidoperationexception_on_empty_peek
$s = [System.Collections.Generic.Stack[int]]::new()
$caught = $false
try {
    $x = $s.Peek()
} catch [System.InvalidOperationException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Peek on empty stack must throw InvalidOperationException"
    exit 1
}
Write-Host "PASS"
exit 0
