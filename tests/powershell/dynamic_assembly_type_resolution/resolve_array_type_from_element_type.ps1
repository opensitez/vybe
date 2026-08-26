# vybe-test: powershell/dynamic_assembly_type_resolution/resolve_array_type_from_element_type
$elemType = [type]"System.String"
$arrType = $elemType.MakeArrayType()
$arr = [System.Array]::CreateInstance($elemType, 3)
$arr.SetValue("A", 0)
$arr.SetValue("B", 1)
$arr.SetValue("C", 2)
if ($arr.Length -ne 3 -or $arr.GetValue(1) -ne "B") {
    Write-Host "FAIL: Dynamic array type creation failed"
    exit 1
}
Write-Host "PASS"
exit 0
