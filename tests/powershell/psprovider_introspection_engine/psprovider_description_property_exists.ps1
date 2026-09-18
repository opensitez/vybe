# vybe-test: powershell/psprovider_introspection_engine/psprovider_description_property_exists
# The Description property exists on ProviderInfo instances
$fs = Get-PSProvider FileSystem

if ($null -eq $fs.PSObject.Properties["Description"]) {
    Write-Host "FAIL: Description property missing from ProviderInfo"
    exit 1
}

Write-Host "PASS"
exit 0
