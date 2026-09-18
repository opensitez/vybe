# vybe-test: powershell/psdrive_management_cmdlets/psdrive_nonexistent_drive_returns_null_silently
# Querying a non-existent drive name with -ErrorAction SilentlyContinue returns $null
$missing = Get-PSDrive -Name "NonExistentDrive_99988" -ErrorAction SilentlyContinue

if ($null -ne $missing) {
    Write-Host "FAIL: expected `$null for non-existent drive, got: $missing"
    exit 1
}

Write-Host "PASS"
exit 0
