# vybe-test: powershell/dynamic_method_invocations_by_string/invoke_string_method_by_variable_name
$method = "ToUpper"
$str = "hello"
$res = $str.$method()
if ($res -ne "HELLO") {
    Write-Host "FAIL: Dynamic method invocation by variable failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
