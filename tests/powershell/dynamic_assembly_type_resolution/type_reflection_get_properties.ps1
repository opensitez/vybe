# vybe-test: powershell/dynamic_assembly_type_resolution/type_reflection_get_properties
$t = [type]"System.Version"
$props = @($t.GetProperties() | ForEach-Object { $_.Name })
if (-not ($props -contains "Major") -or -not ($props -contains "Minor")) {
    Write-Host "FAIL: GetProperties reflection check failed"
    exit 1
}
Write-Host "PASS"
exit 0
