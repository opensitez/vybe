# vybe-test: powershell/using_variable_scope/using_variable_closure_alternative
$limit = 15
$sb = { 10..20 | Where-Object { $_ -gt $using:limit } }
$res = @(&$sb)
if ($res.Count -ne 5 -or $res[0] -ne 16 -or $res[4] -ne 20) {
    Write-Host "FAIL: using variable filter expected 16..20, got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
