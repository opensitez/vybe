# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_clear_and_repopulate
$obj = [pscustomobject]@{ X = 10 }
$obj.PSObject.TypeNames.Clear()
$obj.PSObject.TypeNames.Add("CleanType")
if ($obj.PSObject.TypeNames.Count -ne 1 -or $obj.PSObject.TypeNames[0] -ne "CleanType") {
    Write-Host "FAIL: PSTypeNames Clear and repopulate failed"
    exit 1
}
Write-Host "PASS"
exit 0
