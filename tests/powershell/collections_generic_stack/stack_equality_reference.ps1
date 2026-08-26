# vybe-test: powershell/collections_generic_stack/stack_equality_reference
$s1 = [System.Collections.Generic.Stack[int]]::new()
$s2 = $s1
if ($s1 -ne $s2) {
    Write-Host "FAIL: Stack reference equality failed"
    exit 1
}
Write-Host "PASS"
exit 0
