# vybe-test: powershell/psprovider_introspection_engine/psprovider_filesystem_capabilities_flags
# The FileSystem provider exhibits Filter and ShouldProcess capability flags
$fs = Get-PSProvider FileSystem

$hasFilter = [bool]($fs.Capabilities -band [System.Management.Automation.Provider.ProviderCapabilities]::Filter)
$hasShouldProcess = [bool]($fs.Capabilities -band [System.Management.Automation.Provider.ProviderCapabilities]::ShouldProcess)

if (-not $hasFilter) {
    Write-Host "FAIL: FileSystem provider missing Filter capability"
    exit 1
}

if (-not $hasShouldProcess) {
    Write-Host "FAIL: FileSystem provider missing ShouldProcess capability"
    exit 1
}

Write-Host "PASS"
exit 0
