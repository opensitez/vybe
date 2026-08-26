# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_type_info_inspection
$obj = [pscustomobject]@{ Id = 1 }
$obj.PSObject.TypeNames.Insert(0, "My.Custom.Type")
if ($obj.PSObject.TypeNames[0] -ne "My.Custom.Type") {
    Write-Host "FAIL: TypeNames hierarchy insertion failed"
    exit 1
}
Write-Host "PASS"
exit 0
