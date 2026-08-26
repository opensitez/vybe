# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_multiple_inheritance_simulation
$obj = [pscustomobject]@{ Id = 100 }
$obj.PSObject.TypeNames.Insert(0, "App.Entity")
$obj.PSObject.TypeNames.Insert(0, "App.UserEntity")
if ($obj.PSObject.TypeNames[0] -ne "App.UserEntity" -or $obj.PSObject.TypeNames[1] -ne "App.Entity") {
    Write-Host "FAIL: Multi-level PSTypeNames hierarchy failed"
    exit 1
}
Write-Host "PASS"
exit 0
