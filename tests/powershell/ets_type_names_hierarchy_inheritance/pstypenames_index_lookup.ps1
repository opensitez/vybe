# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_index_lookup
$obj = [pscustomobject]@{ X = 1 }
$obj.PSObject.TypeNames.Insert(0, "FirstType")
$idx = $obj.PSObject.TypeNames.IndexOf("FirstType")
if ($idx -ne 0) {
    Write-Host "FAIL: PSTypeNames IndexOf failed, got $idx"
    exit 1
}
Write-Host "PASS"
exit 0
