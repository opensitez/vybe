# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_on_array_instance
$arr = @(1, 2, 3)
$names = @($arr.PSObject.TypeNames)
if ($names[0] -ne "System.Object[]") {
    Write-Host "FAIL: Array PSTypeNames failed, got '$($names[0])'"
    exit 1
}
Write-Host "PASS"
exit 0
