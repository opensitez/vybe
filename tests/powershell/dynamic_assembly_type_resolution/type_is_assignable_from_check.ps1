# vybe-test: powershell/dynamic_assembly_type_resolution/type_is_assignable_from_check
$baseType = [System.Collections.IEnumerable]
$derivedType = [type]"System.Collections.Generic.List[string]"
if (-not $baseType.IsAssignableFrom($derivedType)) {
    Write-Host "FAIL: IsAssignableFrom check failed"
    exit 1
}
Write-Host "PASS"
exit 0
