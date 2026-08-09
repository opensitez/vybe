# vybe-test: powershell/pstypenames/pstypenames_hierarchy
$obj = [pscustomobject]@{ Val = 10 }
$obj.psobject.TypeNames.Insert(0, "BaseType")
$obj.psobject.TypeNames.Insert(0, "DerivedType")
if ($obj.psobject.TypeNames[0] -ne "DerivedType" -or $obj.psobject.TypeNames[1] -ne "BaseType") {
    Write-Host "FAIL: TypeNames hierarchy expected DerivedType, BaseType"
    exit 1
}
Write-Host "PASS"
exit 0
