# vybe-test: powershell/using_variable_scope/using_variable_multiple_vars
$v1 = 10
$v2 = 20
$sb = { $using:v1 + $using:v2 }
$res = &$sb
if ($res -ne 30) {
    Write-Host "FAIL: multiple \$using: variables expected 30, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
