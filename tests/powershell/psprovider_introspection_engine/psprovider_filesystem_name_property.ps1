# vybe-test: powershell/psprovider_introspection_engine/psprovider_filesystem_name_property
# Get-PSProvider FileSystem returns ProviderInfo with Name 'FileSystem'
$fs = Get-PSProvider FileSystem

if ($null -eq $fs) {
    Write-Host "FAIL: FileSystem provider was not returned"
    exit 1
}

if ($fs.Name -ne "FileSystem") {
    Write-Host "FAIL: expected provider Name 'FileSystem', got: '$($fs.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
