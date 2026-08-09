# vybe-test: powershell/using_variable_scope/using_variable_in_threadjob
$threadData = "ThreadValue"
$sb = { $using:threadData.ToUpper() }
$res = &$sb
if ($res -ne "THREADVALUE") {
    Write-Host "FAIL: thread job using variable method call expected THREADVALUE, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
