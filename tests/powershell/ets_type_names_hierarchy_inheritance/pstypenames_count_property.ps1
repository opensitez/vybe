# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_count_property
$obj = [pscustomobject]@{ X = 1 }
$initialCount = $obj.PSObject.TypeNames.Count
$obj.PSObject.TypeNames.Insert(0, "Extra")
if ($obj.PSObject.TypeNames.Count -ne ($initialCount + 1)) {
    Write-Host "FAIL: PSTypeNames Count update failed"
    exit 1
}
Write-Host "PASS"
exit 0
