# vybe-test: powershell/resolve_path_cmdlet/resolve_path_nonexistent_path_throws_itemnotfound
# Resolving a non-existent path throws an ItemNotFoundException
$threwExpected = $false
$caughtEx = $null

try {
    Resolve-Path "nonexistent_dir_unique_id_332211" -ErrorAction Stop
} catch {
    $threwExpected = $true
    $caughtEx = $_
}

if (-not $threwExpected) {
    Write-Host "FAIL: Resolve-Path on nonexistent path did not throw"
    exit 1
}

if (-not ($caughtEx.Exception -is [System.Management.Automation.ItemNotFoundException])) {
    Write-Host "FAIL: expected ItemNotFoundException, got: $($caughtEx.Exception.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
