# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_returning_boolean
$str = "abc"
$m = "Contains"
$hasB = $str.$m("b")
$hasZ = $str.$m("z")
if (-not $hasB -or $hasZ) {
    Write-Host "FAIL: Dynamic method returning boolean failed"
    exit 1
}
Write-Host "PASS"
exit 0
