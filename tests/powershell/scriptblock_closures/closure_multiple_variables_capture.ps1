# vybe-test: powershell/scriptblock_closures/closure_multiple_variables_capture
$a = 10
$b = 20
$sb = { $a + $b }.GetClosure()
$a = 100
$b = 200
$res = &$sb
if ($res -ne 30) {
    Write-Host "FAIL: GetClosure multiple variables expected (10+20)=30, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
