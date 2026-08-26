# vybe-test: powershell/dynamic_assembly_type_resolution/resolve_generic_type_by_name_string
$typeStr = "System.Collections.Generic.List[string]"
$type = [type]$typeStr
$inst = [Activator]::CreateInstance($type)
$inst.Add("hello")
if ($inst.Count -ne 1 -or $inst[0] -ne "hello") {
    Write-Host "FAIL: Generic type resolution by string failed"
    exit 1
}
Write-Host "PASS"
exit 0
