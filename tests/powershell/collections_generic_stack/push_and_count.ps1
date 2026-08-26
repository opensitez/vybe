# vybe-test: powershell/collections_generic_stack/push_and_count
$s = [System.Collections.Generic.Stack[string]]::new()
$s.Push("bottom")
$s.Push("top")
if ($s.Count -ne 2) {
    Write-Host "FAIL: Stack Push Count mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
