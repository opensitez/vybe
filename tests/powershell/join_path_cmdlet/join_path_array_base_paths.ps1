# vybe-test: powershell/join_path_cmdlet/join_path_array_base_paths
# Passing an array to the -Path parameter evaluates the join across each base path
$res = Join-Path -Path @("usr", "etc") -ChildPath "default"

if ($res.Count -ne 2) {
    Write-Host "FAIL: expected 2 joined paths, got $($res.Count)"
    exit 1
}

if (-not ($res[0] -match "^usr[/|\\]default$" -and $res[1] -match "^etc[/|\\]default$")) {
    Write-Host "FAIL: array base paths join mismatch: @($($res -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
