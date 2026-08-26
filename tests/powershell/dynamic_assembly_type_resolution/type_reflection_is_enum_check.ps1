# vybe-test: powershell/dynamic_assembly_type_resolution/type_reflection_is_enum_check
$tDay = [type]"System.DayOfWeek"
$tInt = [type]"int"
if (-not $tDay.IsEnum -or $tInt.IsEnum) {
    Write-Host "FAIL: IsEnum check failed"
    exit 1
}
Write-Host "PASS"
exit 0
