# vybe-test: powershell/psprovider_introspection_engine/psprovider_environment_drives_collection
# Environment provider exposes the 'Env' PSDrive in its Drives collection
$envP = Get-PSProvider Environment

if ($envP.Name -ne "Environment") {
    Write-Host "FAIL: expected provider Name 'Environment', got: '$($envP.Name)'"
    exit 1
}

$driveNames = @($envP.Drives | ForEach-Object { $_.Name })
if ($driveNames -notcontains "Env") {
    Write-Host "FAIL: 'Env' drive missing from Environment provider Drives: @($($driveNames -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
