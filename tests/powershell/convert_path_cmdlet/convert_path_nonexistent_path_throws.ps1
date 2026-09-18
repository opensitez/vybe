# vybe-test: powershell/convert_path_cmdlet/convert_path_nonexistent_path_throws
# Convert-Path throws an ItemNotFoundException when attempting to resolve a nonexistent path
$threwExpected = $false
$caughtEx = $null

try {
    Convert-Path "nonexistent_dir_unique_id_998811" -ErrorAction Stop
} catch {
    $threwExpected = $true
    $caughtEx = $_
}

if (-not $threwExpected) {
    Write-Host "FAIL: Convert-Path on nonexistent path did not throw"
    exit 1
}

if (-not ($caughtEx.Exception -is [System.Management.Automation.ItemNotFoundException])) {
    Write-Host "FAIL: expected ItemNotFoundException, got: $($caughtEx.Exception.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
