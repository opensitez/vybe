# vybe-test: powershell/psdrive_management_cmdlets/psdrive_new_filesystem_drive_registration
# New-PSDrive mounts a new filesystem drive pointing to a directory root
$drive = New-PSDrive -Name CustomFSDrive -PSProvider FileSystem -Root $pwd.Path

if ($null -eq $drive) {
    Write-Host "FAIL: New-PSDrive returned `$null"
    exit 1
}

if ($drive.Name -ne "CustomFSDrive") {
    Write-Host "FAIL: expected drive name 'CustomFSDrive', got: '$($drive.Name)'"
    exit 1
}

$retrieved = Get-PSDrive -Name CustomFSDrive
Remove-PSDrive -Name CustomFSDrive

if ($null -eq $retrieved) {
    Write-Host "FAIL: created drive could not be retrieved by Get-PSDrive"
    exit 1
}

Write-Host "PASS"
exit 0
