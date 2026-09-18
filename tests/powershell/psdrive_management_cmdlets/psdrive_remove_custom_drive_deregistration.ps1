# vybe-test: powershell/psdrive_management_cmdlets/psdrive_remove_custom_drive_deregistration
# Remove-PSDrive removes a mounted PSDrive so that Get-PSDrive no longer finds it
$drive = New-PSDrive -Name RemovableDrive -PSProvider FileSystem -Root $pwd.Path
Remove-PSDrive -Name RemovableDrive

$found = Get-PSDrive -Name RemovableDrive -ErrorAction SilentlyContinue

if ($null -ne $found) {
    Write-Host "FAIL: drive was still returned by Get-PSDrive after Remove-PSDrive"
    exit 1
}

Write-Host "PASS"
exit 0
