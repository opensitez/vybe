# vybe-test: powershell/ets_type_names_hierarchy_inheritance/pstypenames_on_dotnet_type
$dt = [datetime]::UtcNow
$names = @($dt.PSObject.TypeNames)
if ($names[0] -ne "System.DateTime") {
    Write-Host "FAIL: .NET type PSTypeNames check failed"
    exit 1
}
Write-Host "PASS"
exit 0
