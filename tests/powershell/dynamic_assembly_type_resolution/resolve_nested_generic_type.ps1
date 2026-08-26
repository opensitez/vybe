# vybe-test: powershell/dynamic_assembly_type_resolution/resolve_nested_generic_type
$typeStr = "System.Collections.Generic.List[int]"
$type = [type]$typeStr
$inst = [Activator]::CreateInstance($type)
$inst.Add(42)
if ($inst.Count -ne 1 -or $inst[0] -ne 42) {
    Write-Host "FAIL: Type resolution failed"
    exit 1
}
Write-Host "PASS"
exit 0
