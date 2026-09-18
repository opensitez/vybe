# vybe-test: powershell/psprovider_introspection_engine/psprovider_alias_drives_collection
# Alias provider exposes the 'Alias' PSDrive in its Drives collection
$aliasP = Get-PSProvider Alias

if ($aliasP.Name -ne "Alias") {
    Write-Host "FAIL: expected provider Name 'Alias', got: '$($aliasP.Name)'"
    exit 1
}

$driveNames = @($aliasP.Drives | ForEach-Object { $_.Name })
if ($driveNames -notcontains "Alias") {
    Write-Host "FAIL: 'Alias' drive missing from Alias provider Drives"
    exit 1
}

Write-Host "PASS"
exit 0
