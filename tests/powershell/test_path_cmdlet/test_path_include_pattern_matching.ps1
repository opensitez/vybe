# vybe-test: powershell/test_path_cmdlet/test_path_include_pattern_matching
# -Include filters path targets against an inclusion glob pattern
$res = Test-Path "tests/powershell" -Include "*powershell*"

if ($res -ne $true) {
    Write-Host "FAIL: Test-Path with -Include pattern expected `$true, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
