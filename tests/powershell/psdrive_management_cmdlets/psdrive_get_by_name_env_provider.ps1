# vybe-test: powershell/psdrive_management_cmdlets/psdrive_get_by_name_env_provider
# Get-PSDrive -Name Env queries the Environment provider drive
$envDrive = Get-PSDrive -Name Env

if ($null -eq $envDrive) {
    Write-Host "FAIL: Env drive was not returned"
    exit 1
}

if ($envDrive.Name -ne "Env") {
    Write-Host "FAIL: expected drive Name 'Env', got: '$($envDrive.Name)'"
    exit 1
}

if ($envDrive.Provider.Name -ne "Environment") {
    Write-Host "FAIL: expected provider 'Environment', got: '$($envDrive.Provider.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
