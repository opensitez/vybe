# vybe-test: powershell/scriptblock_closures/closure_basic_variable_capture
$outerVal = 42
$sb = { $outerVal }.GetClosure()
$outerVal = 99
$res = &$sb
if ($res -ne 42) {
    Write-Host "FAIL: GetClosure expected captured snapshot 42, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
