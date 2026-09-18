# vybe-test: powershell/psprovider_introspection_engine/psprovider_removing_psdrive_updates_provider_drives
# Removing a PSDrive immediately removes it from the provider's Drives collection
$drive = New-PSDrive -Name ProvCleanDrive -PSProvider FileSystem -Root $pwd.Path
Remove-PSDrive -Name ProvCleanDrive

$fs = Get-PSProvider FileSystem
$names = @($fs.Drives | ForEach-Object { $_.Name })

if ($names -contains "ProvCleanDrive") {
    Write-Host "FAIL: removed drive ProvCleanDrive still lingered in provider Drives collection"
    exit 1
}

Write-Host "PASS"
exit 0
