# vybe-test: powershell/resolve_path_cmdlet/resolve_path_array_argument_parameter
# Passing an array of paths to the -Path parameter returns an array of PathInfo objects
$results = Resolve-Path -Path @(".", "tests")

if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 results, got $($results.Count)"
    exit 1
}

if (-not ($results[0] -is [System.Management.Automation.PathInfo] -and $results[1] -is [System.Management.Automation.PathInfo])) {
    Write-Host "FAIL: elements in returned array are not PathInfo objects"
    exit 1
}

Write-Host "PASS"
exit 0
