# vybe-test: powershell/psdrive_management_cmdlets/psdrive_get_multiple_drives_by_name_array
# Get-PSDrive accepts an array of drive names to query multiple drives simultaneously
$drives = @(Get-PSDrive -Name Env, Variable)

if ($drives.Count -ne 2) {
    Write-Host "FAIL: expected 2 drives from array query, got $($drives.Count)"
    exit 1
}

$names = @($drives | ForEach-Object { $_.Name })
if ($names -notcontains "Env" -or $names -notcontains "Variable") {
    Write-Host "FAIL: expected 'Env' and 'Variable' in results: @($($names -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
