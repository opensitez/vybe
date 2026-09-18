# vybe-test: powershell/psprovider_introspection_engine/psprovider_home_property_exists
# The Home property exists on ProviderInfo instances representing provider home location
$fs = Get-PSProvider FileSystem

if ($null -eq $fs.PSObject.Properties["Home"]) {
    Write-Host "FAIL: Home property missing from FileSystem ProviderInfo"
    exit 1
}

Write-Host "PASS"
exit 0
