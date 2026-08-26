# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_on_derived_class_contains_base_type
class BaseEntity {}
class DerivedEntity : BaseEntity {}
$de = [DerivedEntity]::new()
$names = @($de.PSObject.TypeNames)
if ($names[0] -ne "DerivedEntity" -or $names[1] -ne "BaseEntity") {
    Write-Host "FAIL: Derived class PSTypeNames hierarchy failed"
    exit 1
}
Write-Host "PASS"
exit 0
