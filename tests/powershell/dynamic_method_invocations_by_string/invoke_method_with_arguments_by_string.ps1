# vybe-test: powershell/dynamic_method_invocations_by_string/invoke_method_with_arguments_by_string
$method = "Substring"
$str = "PowerShell"
$res = $str.$method(0, 5)
if ($res -ne "Power") {
    Write-Host "FAIL: Dynamic method invocation with args failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
