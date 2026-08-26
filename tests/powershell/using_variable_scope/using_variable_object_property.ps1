# vybe-test: powershell/using_variable_scope/using_variable_object_property
$localNum = 42
$job = Start-ThreadJob -ScriptBlock { $using:localNum }
$res = Receive-Job $job -Wait -AutoRemoveJob
if ($res -eq 42) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
