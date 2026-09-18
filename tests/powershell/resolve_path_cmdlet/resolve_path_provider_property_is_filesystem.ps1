# vybe-test: powershell/resolve_path_cmdlet/resolve_path_provider_property_is_filesystem
# The .Provider property of the returned PathInfo identifies the PowerShell provider
$info = Resolve-Path "."

if ($info.Provider.Name -ne "FileSystem") {
    Write-Host "FAIL: expected provider 'FileSystem', got: '$($info.Provider.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
