# vybe-test: powershell/convert_path_cmdlet/convert_path_multiple_wildcard_matches
# Wildcards matching multiple filesystem entries return an array of resolved paths
$results = Convert-Path "tests/*"

if ($null -eq $results) {
    Write-Host "FAIL: Convert-Path wildcard returned `$null"
    exit 1
}

if ($results.Count -lt 2) {
    Write-Host "FAIL: expected multiple wildcard results in tests/*, got count $($results.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
