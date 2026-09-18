# vybe-test: powershell/join_path_cmdlet/join_path_resolve_nonexistent_throws
# -Resolve throws an ItemNotFoundException when the target path does not exist
$threwExpected = $false
$caughtEx = $null

try {
    Join-Path "." "nonexistent_dir_random_unique_9988" -Resolve -ErrorAction Stop
} catch {
    $threwExpected = $true
    $caughtEx = $_
}

if (-not $threwExpected) {
    Write-Host "FAIL: Join-Path -Resolve on nonexistent path did not throw"
    exit 1
}

if (-not ($caughtEx.Exception -is [System.Management.Automation.ItemNotFoundException])) {
    Write-Host "FAIL: expected ItemNotFoundException, got: $($caughtEx.Exception.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
