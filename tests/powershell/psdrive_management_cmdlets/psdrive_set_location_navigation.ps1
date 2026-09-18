# vybe-test: powershell/psdrive_management_cmdlets/psdrive_set_location_navigation
# Set-Location navigates into a custom PSDrive using the 'DriveName:' syntax
$drive = New-PSDrive -Name NavTargetDrive -PSProvider FileSystem -Root $pwd.Path
$origLoc = $pwd.Path

Set-Location NavTargetDrive:
$curLoc = Get-Location
Set-Location $origLoc
Remove-PSDrive -Name NavTargetDrive

if ($curLoc.Drive.Name -ne "NavTargetDrive") {
    Write-Host "FAIL: active drive after Set-Location expected 'NavTargetDrive', got: '$($curLoc.Drive.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
