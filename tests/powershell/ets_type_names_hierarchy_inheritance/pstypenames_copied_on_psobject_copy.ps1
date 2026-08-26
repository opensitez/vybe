# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_copied_on_psobject_copy
$obj = [pscustomobject]@{ Tag = "A" }
$obj.PSObject.TypeNames.Insert(0, "SpecialType")
$copy = $obj.PSObject.Copy()
if ($copy.PSObject.TypeNames[0] -ne "SpecialType") {
    Write-Host "FAIL: PSTypeNames preservation on Copy() failed"
    exit 1
}
Write-Host "PASS"
exit 0
