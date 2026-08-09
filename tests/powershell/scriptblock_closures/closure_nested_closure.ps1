# vybe-test: powershell/scriptblock_closures/closure_nested_closure
$n1 = 5
$outerSb = {
    $n2 = 10
    return { $n1 + $n2 }.GetClosure()
}.GetClosure()
$innerSb = &$outerSb
$res = &$innerSb
if ($res -ne 15) {
    Write-Host "FAIL: nested closure evaluation expected (5+10)=15, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
