# vybe-test: powershell/test_path_cmdlet/test_path_hidden_dot_directory
# Test-Path detects hidden dot-directories such as .git
$res = Test-Path ".git"

if ($res -ne $true) {
    Write-Host "FAIL: expected `$true for existing hidden directory '.git', got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
