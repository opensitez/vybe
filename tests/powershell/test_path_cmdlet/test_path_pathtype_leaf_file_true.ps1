# vybe-test: powershell/test_path_cmdlet/test_path_pathtype_leaf_file_true
# Test-Path with -PathType Leaf returns $true for existing files
$res = Test-Path "Cargo.toml" -PathType Leaf

if ($res -ne $true) {
    Write-Host "FAIL: expected `$true for file tested with -PathType Leaf, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
