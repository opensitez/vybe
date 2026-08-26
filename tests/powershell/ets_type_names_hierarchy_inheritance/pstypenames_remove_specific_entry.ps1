# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_remove_specific_entry
$obj = [pscustomobject]@{ X = 1 }
$obj.PSObject.TypeNames.Insert(0, "TempType")
$obj.PSObject.TypeNames.Remove("TempType")
if ($obj.PSObject.TypeNames -contains "TempType") {
    Write-Host "FAIL: PSTypeNames Remove failed"
    exit 1
}
Write-Host "PASS"
exit 0
