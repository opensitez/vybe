# vybe-test: powershell/test_path_cmdlet/test_path_wildcard_matching_returns_true
# Wildcard path matches return $true when matching elements exist
$res = Test-Path "tests/pow*"

if ($res -ne $true) {
    Write-Host "FAIL: wildcard matching expected `$true, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
