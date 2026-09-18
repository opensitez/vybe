# vybe-test: powershell/psprovider_introspection_engine/psprovider_variable_drives_collection
# Variable provider exposes the 'Variable' PSDrive in its Drives collection
$varP = Get-PSProvider Variable

if ($varP.Name -ne "Variable") {
    Write-Host "FAIL: expected provider Name 'Variable', got: '$($varP.Name)'"
    exit 1
}

$driveNames = @($varP.Drives | ForEach-Object { $_.Name })
if ($driveNames -notcontains "Variable") {
    Write-Host "FAIL: 'Variable' drive missing from Variable provider Drives"
    exit 1
}

Write-Host "PASS"
exit 0
