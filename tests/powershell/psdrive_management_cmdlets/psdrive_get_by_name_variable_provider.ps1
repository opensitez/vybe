# vybe-test: powershell/psdrive_management_cmdlets/psdrive_get_by_name_variable_provider
# Get-PSDrive -Name Variable queries the Variable provider drive
$varDrive = Get-PSDrive -Name Variable

if ($null -eq $varDrive) {
    Write-Host "FAIL: Variable drive was not returned"
    exit 1
}

if ($varDrive.Name -ne "Variable" -or $varDrive.Provider.Name -ne "Variable") {
    Write-Host "FAIL: drive or provider mismatch: Name='$($varDrive.Name)', Provider='$($varDrive.Provider.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
