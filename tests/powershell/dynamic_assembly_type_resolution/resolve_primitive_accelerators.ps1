# vybe-test: powershell/dynamic_assembly_type_resolution/resolve_primitive_accelerators
$tInt = [type]"int"
$tStr = [type]"string"
$tBool = [type]"bool"
$tLong = [type]"long"
if ($tInt -ne [int] -or $tStr -ne [string] -or $tBool -ne [bool] -or $tLong -ne [int64]) {
    Write-Host "FAIL: Primitive type accelerator resolution failed"
    exit 1
}
Write-Host "PASS"
exit 0
