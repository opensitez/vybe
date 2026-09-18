# vybe-test: powershell/psdrive_management_cmdlets/psdrive_get_by_name_function_provider
# Get-PSDrive -Name Function queries the Function provider drive
$funcDrive = Get-PSDrive -Name Function

if ($null -eq $funcDrive) {
    Write-Host "FAIL: Function drive was not returned"
    exit 1
}

if ($funcDrive.Name -ne "Function" -or $funcDrive.Provider.Name -ne "Function") {
    Write-Host "FAIL: drive or provider mismatch: Name='$($funcDrive.Name)', Provider='$($funcDrive.Provider.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
