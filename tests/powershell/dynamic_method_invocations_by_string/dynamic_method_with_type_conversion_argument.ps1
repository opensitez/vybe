# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_with_type_conversion_argument
$method = "Parse"
$res = [int]::$method("12345")
if ($res -ne 12345) {
    Write-Host "FAIL: Dynamic type parse static method failed"
    exit 1
}
Write-Host "PASS"
exit 0
