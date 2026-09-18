# vybe-test: powershell/join_path_cmdlet/join_path_resolve_wildcard_matching
# -Resolve with wildcard expression matches and resolves existing matching directories
$resolved = Join-Path "tests" "power*" -Resolve

if ($null -eq $resolved) {
    Write-Host "FAIL: Join-Path -Resolve with wildcard returned `$null"
    exit 1
}

if ($resolved -notmatch "powershell") {
    Write-Host "FAIL: resolved wildcard path did not match 'powershell': '$resolved'"
    exit 1
}

Write-Host "PASS"
exit 0
