# vybe-test: powershell/using_variable_scope/using_variable_loop_iteration
$multiplier = 7
$sb = { foreach ($i in 1..3) { $i * $using:multiplier } }
$res = @(&$sb)
if ($res[0] -ne 7 -or $res[1] -ne 14 -or $res[2] -ne 21) {
    Write-Host "FAIL: loop iteration with \$using:multiplier expected 7, 14, 21"
    exit 1
}
Write-Host "PASS"
exit 0
