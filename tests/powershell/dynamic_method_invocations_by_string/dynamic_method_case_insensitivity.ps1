# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_case_insensitivity
$str = "abc"
$m = "touPPer"
$res = $str.$m()
if ($res -ne "ABC") {
    Write-Host "FAIL: Dynamic method case-insensitivity failed"
    exit 1
}
Write-Host "PASS"
exit 0
