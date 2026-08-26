# vybe-test: powershell/scriptblock_closures/closure_null_capture
$nullVal = $null
$sb = { $nullVal }.GetNewClosure()
$res = &$sb
if ($res -ne $null) {
    Write-Host "FAIL: null capture in closure expected null, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
