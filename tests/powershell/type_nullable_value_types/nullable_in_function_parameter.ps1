# vybe-test: powershell/type_nullable_value_types/nullable_in_function_parameter
function Test-NullableParam {
    param([System.Nullable[int]]$val)
    if ($null -ne $val) { return $val * 2 }
    return -1
}
$r1 = Test-NullableParam -val 5
$r2 = Test-NullableParam -val $null
if ($r1 -ne 10 -or $r2 -ne -1) {
    Write-Host "FAIL: Function nullable parameter dispatch failed"
    exit 1
}
Write-Host "PASS"
exit 0
