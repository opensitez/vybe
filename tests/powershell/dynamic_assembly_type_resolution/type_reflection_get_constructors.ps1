# vybe-test: powershell/dynamic_assembly_type_resolution/type_reflection_get_constructors
$t = [type]"System.Text.StringBuilder"
$constructors = @($t.GetConstructors())
if ($constructors.Length -le 1) {
    Write-Host "FAIL: GetConstructors reflection check failed, count=$($constructors.Length)"
    exit 1
}
Write-Host "PASS"
exit 0
