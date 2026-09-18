# vybe-test: powershell/resolve_path_cmdlet/resolve_path_pipeline_streaming
$paths = @(".", "tests")

# Piping paths into Resolve-Path emits PathInfo objects for each input
$results = @($paths | Resolve-Path)

if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 resolved items, got $($results.Count)"
    exit 1
}

if ($results[0].Path -ne $PWD.Path -or ($results[1].Path -notmatch "tests$")) {
    Write-Host "FAIL: pipeline resolved items mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
