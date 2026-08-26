# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_with_array_argument
$m = "Join"
$res = [string]::$m("-", @("a", "b", "c"))
if ($res -ne "a-b-c") {
    Write-Host "FAIL: Dynamic static method with array argument failed"
    exit 1
}
Write-Host "PASS"
exit 0
