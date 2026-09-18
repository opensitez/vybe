# vybe-test: powershell/psdrive_management_cmdlets/psdrive_new_with_description_property
# New-PSDrive assigns the -Description parameter to the PSDriveInfo Description property
$drive = New-PSDrive -Name DescDriveTest -PSProvider FileSystem -Root $pwd.Path -Description "Project Workspace Mount"
$desc = $drive.Description
Remove-PSDrive -Name DescDriveTest

if ($desc -ne "Project Workspace Mount") {
    Write-Host "FAIL: Description mismatch, expected 'Project Workspace Mount', got: '$desc'"
    exit 1
}

Write-Host "PASS"
exit 0
