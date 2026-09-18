# vybe-test: powershell/resolve_path_cmdlet/resolve_path_multiple_wildcard_matches
# Wildcards matching multiple items return a collection of PathInfo objects
$results = Resolve-Path "tests/*"

if ($null -eq $results) {
    Write-Host "FAIL: wildcard resolution returned `$null"
    exit 1
}

if ($results.Count -lt 2) {
    Write-Host "FAIL: expected multiple items in tests/*, got $($results.Count)"
    exit 1
}

if (-not ($results[0] -is [System.Management.Automation.PathInfo])) {
    Write-Host "FAIL: wildcard results are not PathInfo objects"
    exit 1
}

Write-Host "PASS"
exit 0
