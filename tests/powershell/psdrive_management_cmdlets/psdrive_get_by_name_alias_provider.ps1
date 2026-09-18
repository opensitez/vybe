# vybe-test: powershell/psdrive_management_cmdlets/psdrive_get_by_name_alias_provider
# Get-PSDrive -Name Alias queries the Alias provider drive
$aliasDrive = Get-PSDrive -Name Alias

if ($null -eq $aliasDrive) {
    Write-Host "FAIL: Alias drive was not returned"
    exit 1
}

if ($aliasDrive.Name -ne "Alias" -or $aliasDrive.Provider.Name -ne "Alias") {
    Write-Host "FAIL: drive or provider mismatch: Name='$($aliasDrive.Name)', Provider='$($aliasDrive.Provider.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
