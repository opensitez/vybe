# vybe-test: powershell/dynamic_assembly_type_resolution/type_reflection_is_value_type_check
$tInt = [type]"int"
$tStr = [type]"string"
if (-not $tInt.IsValueType -or $tStr.IsValueType) {
    Write-Host "FAIL: IsValueType check failed"
    exit 1
}
Write-Host "PASS"
exit 0
