# vybe-test: powershell/scriptblock_closures/closure_with_parameters
$prefix = "LOG"
$sb = { param([string]$msg) "$prefix: $msg" }.GetClosure()
$res = &$sb "SystemStarted"
if ($res -ne "LOG: SystemStarted") {
    Write-Host "FAIL: parameterized closure expected 'LOG: SystemStarted', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
