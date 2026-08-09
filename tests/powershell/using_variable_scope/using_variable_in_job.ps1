# vybe-test: powershell/using_variable_scope/using_variable_in_job
$jobVal = 555
$sb = { $using:jobVal * 2 }
$res = &$sb
if ($res -ne 1110) {
    Write-Host "FAIL: job scriptblock using variable expected 1110, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
