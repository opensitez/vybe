# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_custom_type_name_prepend
$obj = [pscustomobject]@{ Name = "Item1" }
$obj.PSObject.TypeNames.Insert(0, "MyCustom.Type")
if ($obj.PSObject.TypeNames[0] -ne "MyCustom.Type") {
    Write-Host "FAIL: PSTypeNames insert at 0 failed"
    exit 1
}
Write-Host "PASS"
exit 0
