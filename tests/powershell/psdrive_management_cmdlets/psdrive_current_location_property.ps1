# vybe-test: powershell/psdrive_management_cmdlets/psdrive_current_location_property
# PSDriveInfo exposes a .CurrentLocation property reflecting relative directory state
$drive = New-PSDrive -Name CurrLocTestDrive -PSProvider FileSystem -Root $pwd.Path

if ($null -eq $drive.CurrentLocation) {
    Write-Host "FAIL: CurrentLocation property is `$null on newly created PSDrive"
    Remove-PSDrive -Name CurrLocTestDrive
    exit 1
}

Remove-PSDrive -Name CurrLocTestDrive

Write-Host "PASS"
exit 0
