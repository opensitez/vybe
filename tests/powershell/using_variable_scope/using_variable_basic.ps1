# vybe-test: powershell/using_variable_scope/using_variable_basic
$localNum = 42
$sb = { $using:localNum }
$res = &$sb
if ($res -ne 42) {
    Write-Host "FAIL: \$using:localNum basic resolution expected 42, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
