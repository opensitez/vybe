# vybe-test: powershell/dynamic_method_invocations_by_string/invoke_math_static_method_dynamically
$method = "Abs"
$res = [math]::$method(-42)
if ($res -ne 42) {
    Write-Host "FAIL: Dynamic static method invocation failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
