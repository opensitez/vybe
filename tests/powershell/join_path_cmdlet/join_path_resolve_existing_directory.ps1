# vybe-test: powershell/join_path_cmdlet/join_path_resolve_existing_directory
# -Resolve joins the path and resolves it to a full canonical filesystem path
$resolved = Join-Path "." "tests" -Resolve

if ($null -eq $resolved) {
    Write-Host "FAIL: Join-Path -Resolve returned `$null"
    exit 1
}

if (-not [System.IO.Directory]::Exists($resolved)) {
    Write-Host "FAIL: resolved path is not an existing directory: '$resolved'"
    exit 1
}

Write-Host "PASS"
exit 0
