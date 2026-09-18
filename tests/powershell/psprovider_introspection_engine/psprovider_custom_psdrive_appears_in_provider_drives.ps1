# vybe-test: powershell/psprovider_introspection_engine/psprovider_custom_psdrive_appears_in_provider_drives
# Creating a custom PSDrive dynamically registers it in the corresponding provider's Drives collection
$drive = New-PSDrive -Name ProvDynDrive -PSProvider FileSystem -Root $pwd.Path

$fs = Get-PSProvider FileSystem
$names = @($fs.Drives | ForEach-Object { $_.Name })

Remove-PSDrive -Name ProvDynDrive

if ($names -notcontains "ProvDynDrive") {
    Write-Host "FAIL: ProvDynDrive did not appear in FileSystem provider Drives collection"
    exit 1
}

Write-Host "PASS"
exit 0
