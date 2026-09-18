# vybe-test: powershell/test_path_cmdlet/test_path_nonmatching_wildcard_returns_false
# Wildcard patterns that match zero filesystem items return $false
$res = Test-Path "tests/nonexistent_prefix_pattern_xyz_*"

if ($res -ne $false) {
    Write-Host "FAIL: non-matching wildcard expected `$false, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
