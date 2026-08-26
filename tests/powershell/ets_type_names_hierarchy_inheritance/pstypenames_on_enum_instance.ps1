# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_on_enum_instance
enum StateCode { Off; On }
$s = [StateCode]::On
$names = @($s.PSObject.TypeNames)
if ($names[0] -ne "StateCode") {
    Write-Host "FAIL: Enum PSTypeNames failed, got '$($names[0])'"
    exit 1
}
Write-Host "PASS"
exit 0
