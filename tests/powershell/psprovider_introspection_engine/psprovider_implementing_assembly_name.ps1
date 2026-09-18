# vybe-test: powershell/psprovider_introspection_engine/psprovider_implementing_assembly_name
# Core PowerShell providers are implemented within the System.Management.Automation assembly
$fs = Get-PSProvider FileSystem
$asmName = $fs.ImplementingType.Assembly.GetName().Name

if ($asmName -ne "System.Management.Automation") {
    Write-Host "FAIL: expected assembly 'System.Management.Automation', got: '$asmName'"
    exit 1
}

Write-Host "PASS"
exit 0
