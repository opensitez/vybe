# vybe-test: powershell/psprovider_introspection_engine/psprovider_filesystem_implementing_type
# FileSystem provider implementing type is Microsoft.PowerShell.Commands.FileSystemProvider
$fs = Get-PSProvider FileSystem

$expectedTypeName = "Microsoft.PowerShell.Commands.FileSystemProvider"
if ($fs.ImplementingType.FullName -ne $expectedTypeName) {
    Write-Host "FAIL: expected implementing type '$expectedTypeName', got: '$($fs.ImplementingType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
