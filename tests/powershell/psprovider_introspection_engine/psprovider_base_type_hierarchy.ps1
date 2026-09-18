# vybe-test: powershell/psprovider_introspection_engine/psprovider_base_type_hierarchy
# The FileSystem provider derives from NavigationCmdletProvider in the provider inheritance tree
$fs = Get-PSProvider FileSystem
$baseTypeName = $fs.ImplementingType.BaseType.FullName

if ($baseTypeName -ne "System.Management.Automation.Provider.NavigationCmdletProvider") {
    Write-Host "FAIL: expected NavigationCmdletProvider base, got: '$baseTypeName'"
    exit 1
}

Write-Host "PASS"
exit 0
