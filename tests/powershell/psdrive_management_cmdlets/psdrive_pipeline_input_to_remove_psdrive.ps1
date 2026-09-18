# vybe-test: powershell/psdrive_management_cmdlets/psdrive_pipeline_input_to_remove_psdrive
# Piping PSDriveInfo instances into Remove-PSDrive removes each drive
$drive = New-PSDrive -Name PipeInputDrive -PSProvider FileSystem -Root $pwd.Path

$drive | Remove-PSDrive

$after = Get-PSDrive -Name PipeInputDrive -ErrorAction SilentlyContinue

if ($null -ne $after) {
    Write-Host "FAIL: drive still existed after piping to Remove-PSDrive"
    exit 1
}

Write-Host "PASS"
exit 0
