# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_collection_contains_psobject_and_customobject
$obj = [pscustomobject]@{ A = 1 }
$names = @($obj.PSObject.TypeNames)
if ($names -notcontains "System.Management.Automation.PSCustomObject" -or $names -notcontains "System.Object") {
    Write-Host "FAIL: PSTypeNames default hierarchy check failed"
    exit 1
}
Write-Host "PASS"
exit 0
