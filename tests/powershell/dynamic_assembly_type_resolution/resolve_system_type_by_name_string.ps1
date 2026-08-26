# vybe-test: powershell/dynamic_assembly_type_resolution/resolve_system_type_by_name_string
$typeStr = "System.Text.StringBuilder"
$type = [type]$typeStr
if ($type -ne [System.Text.StringBuilder]) {
    Write-Host "FAIL: Type resolution by string failed, got $type"
    exit 1
}
Write-Host "PASS"
exit 0
