# vybe-test: powershell/psdrive_management_cmdlets/psdrive_root_property_matches_specified_path
# The .Root property of PSDriveInfo preserves the exact root path specified during creation
$targetPath = $pwd.Path
$drive = New-PSDrive -Name RootMatchDrive -PSProvider FileSystem -Root $targetPath
$rootVal = $drive.Root
Remove-PSDrive -Name RootMatchDrive

if ($rootVal -ne $targetPath) {
    Write-Host "FAIL: Root property mismatch, expected '$targetPath', got: '$rootVal'"
    exit 1
}

Write-Host "PASS"
exit 0
