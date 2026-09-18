# vybe-test: powershell/psprovider_introspection_engine/psprovider_capabilities_enum_type
# The Capabilities property is typed as System.Management.Automation.Provider.ProviderCapabilities
$fs = Get-PSProvider FileSystem

$capType = $fs.Capabilities.GetType().FullName

if ($capType -ne "System.Management.Automation.Provider.ProviderCapabilities") {
    Write-Host "FAIL: unexpected Capabilities type, got: '$capType'"
    exit 1
}

Write-Host "PASS"
exit 0
