# vybe-test: powershell/dynamic_assembly_type_resolution/dynamic_type_casting_in_variable_assignment
$typeName = "int"
$type = [type]$typeName
$val = [System.Convert]::ChangeType("456", $type)
if ($val -ne 456 -or $val -isnot [int]) {
    Write-Host "FAIL: ChangeType with dynamic type failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
