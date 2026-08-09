# vybe-test: powershell/using_variable_scope/using_variable_function_call
function Get-Multiplier { return 10 }
$factor = Get-Multiplier
$sb = { 5 * $using:factor }
$res = &$sb
if ($res -ne 50) {
    Write-Host "FAIL: using variable from function return expected 50, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
