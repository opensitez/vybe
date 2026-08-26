# vybe-test: powershell/dynamic_assembly_type_resolution/create_instance_from_dynamically_resolved_type
$typeName = "System.Collections.Generic.HashSet[int]"
$type = [type]$typeName
$hs = [Activator]::CreateInstance($type)
$hs.Add(10)
$hs.Add(20)
if ($hs.Count -ne 2) {
    Write-Host "FAIL: CreateInstance from resolved type failed"
    exit 1
}
Write-Host "PASS"
exit 0
