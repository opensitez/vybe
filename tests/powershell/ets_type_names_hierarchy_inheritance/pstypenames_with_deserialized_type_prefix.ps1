# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_with_deserialized_type_prefix
$obj = [pscustomobject]@{ X = 1 }
$obj.PSObject.TypeNames.Insert(0, "Deserialized.MyNamespace.MyType")
if ($obj.PSObject.TypeNames[0] -ne "Deserialized.MyNamespace.MyType") {
    Write-Host "FAIL: Deserialized PSTypeNames prefix failed"
    exit 1
}
Write-Host "PASS"
exit 0
