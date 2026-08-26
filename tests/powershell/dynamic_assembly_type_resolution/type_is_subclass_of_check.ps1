# vybe-test: powershell/dynamic_assembly_type_resolution/type_is_subclass_of_check
$subType = [type]"System.IO.MemoryStream"
$baseType = [type]"System.IO.Stream"
if (-not $subType.IsSubclassOf($baseType)) {
    Write-Host "FAIL: IsSubclassOf check failed"
    exit 1
}
Write-Host "PASS"
exit 0
