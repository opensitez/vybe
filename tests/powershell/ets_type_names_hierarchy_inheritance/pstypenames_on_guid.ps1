# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_on_guid
$g = [guid]::NewGuid()
$names = @($g.PSObject.TypeNames)
if ($names[0] -ne "System.Guid") {
    Write-Host "FAIL: GUID PSTypeNames check failed"
    exit 1
}
Write-Host "PASS"
exit 0
