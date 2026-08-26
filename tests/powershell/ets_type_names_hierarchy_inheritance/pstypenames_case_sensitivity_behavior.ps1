# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_case_sensitivity_behavior
$obj = [pscustomobject]@{ X = 1 }
$obj.PSObject.TypeNames.Insert(0, "custom.type")
if (-not ($obj.PSObject.TypeNames -contains "custom.type")) {
    Write-Host "FAIL: PSTypeNames case containment failed"
    exit 1
}
Write-Host "PASS"
exit 0
