# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_type_name_reflection
class ReflectionEx : System.Exception {}
$ex = [ReflectionEx]::new()
if ($ex.GetType().Name -ne "ReflectionEx" -or $ex.GetType().BaseType.Name -ne "Exception") {
    Write-Host "FAIL: Custom exception reflection check failed"
    exit 1
}
Write-Host "PASS"
exit 0
