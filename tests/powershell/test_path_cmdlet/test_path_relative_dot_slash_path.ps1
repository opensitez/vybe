# vybe-test: powershell/test_path_cmdlet/test_path_relative_dot_slash_path
# Relative paths with explicit leading './' resolve and test correctly
$res = Test-Path "./Cargo.toml"

if ($res -ne $true) {
    Write-Host "FAIL: expected `$true for './Cargo.toml', got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
