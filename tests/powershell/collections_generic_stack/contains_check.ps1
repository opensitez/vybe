# vybe-test: powershell/collections_generic_stack/contains_check
$s = [System.Collections.Generic.Stack[int]]::new()
$s.Push(100); $s.Push(200)
if (-not $s.Contains(100) -or $s.Contains(999)) {
    Write-Host "FAIL: Stack Contains check failed"
    exit 1
}
Write-Host "PASS"
exit 0
