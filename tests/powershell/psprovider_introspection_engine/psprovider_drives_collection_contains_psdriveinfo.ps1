# vybe-test: powershell/psprovider_introspection_engine/psprovider_drives_collection_contains_psdriveinfo
# The Drives property of ProviderInfo returns a collection of PSDriveInfo objects
$fs = Get-PSProvider FileSystem

if ($fs.Drives.Count -lt 1) {
    Write-Host "FAIL: FileSystem provider has no mounted drives"
    exit 1
}

if (-not ($fs.Drives[0] -is [System.Management.Automation.PSDriveInfo])) {
    Write-Host "FAIL: expected PSDriveInfo in provider Drives, got: $($fs.Drives[0].GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
