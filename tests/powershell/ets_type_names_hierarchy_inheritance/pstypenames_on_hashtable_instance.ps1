# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_on_hashtable_instance
$ht = @{}
$names = @($ht.PSObject.TypeNames)
if ($names[0] -ne "System.Collections.Hashtable") {
    Write-Host "FAIL: Hashtable PSTypeNames failed, got '$($names[0])'"
    exit 1
}
Write-Host "PASS"
exit 0
