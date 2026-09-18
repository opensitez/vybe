# vybe-test: powershell/test_path_cmdlet/test_path_nonexistent_path_returns_false
# Non-existent paths return $false without throwing exceptions
$res = Test-Path "completely_fabricated_path_id_998877"

if ($res -ne $false) {
    Write-Host "FAIL: expected `$false for non-existent path, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
