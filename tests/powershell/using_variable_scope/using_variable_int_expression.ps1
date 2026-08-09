# vybe-test: powershell/using_variable_scope/using_variable_int_expression
$base = 100
$sb = { param($mult) $using:base * $mult }
$res = &$sb 3
if ($res -ne 300) {
    Write-Host "FAIL: \$using:base * mult expected 300, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
