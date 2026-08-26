# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_on_class_instance
class DomainModel {}
$dm = [DomainModel]::new()
if ($dm.PSObject.TypeNames[0] -ne "DomainModel") {
    Write-Host "FAIL: Class instance PSTypeNames check failed, got '$($dm.PSObject.TypeNames[0])'"
    exit 1
}
Write-Host "PASS"
exit 0
