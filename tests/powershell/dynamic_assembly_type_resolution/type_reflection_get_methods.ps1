# vybe-test: powershell/dynamic_assembly_type_resolution/type_reflection_get_methods
$t = [type]"System.Guid"
$methods = @($t.GetMethods() | ForEach-Object { $_.Name })
if (-not ($methods -contains "NewGuid") -or -not ($methods -contains "ToByteArray")) {
    Write-Host "FAIL: GetMethods reflection check failed"
    exit 1
}
Write-Host "PASS"
exit 0
