# vybe-test: powershell/collections_generic_stack/pop_lifo_ordering
$s = [System.Collections.Generic.Stack[int]]::new()
$s.Push(10)
$s.Push(20)
$s.Push(30)
$p1 = $s.Pop() # 30
$p2 = $s.Pop() # 20
if ($p1 -ne 30 -or $p2 -ne 20 -or $s.Count -ne 1) {
    Write-Host "FAIL: Stack Pop LIFO ordering failed"
    exit 1
}
Write-Host "PASS"
exit 0
