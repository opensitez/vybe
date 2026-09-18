# vybe-test: powershell/test_path_cmdlet/test_path_exclude_pattern_matching
# -Exclude filters out matching items, returning $false when the target matches the exclusion pattern
$res = Test-Path "tests/powershell" -Exclude "*powershell*"

if ($res -ne $false) {
    Write-Host "FAIL: Test-Path with -Exclude pattern expected `$false, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
