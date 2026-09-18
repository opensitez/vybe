# vybe-test: powershell/test_path_cmdlet/test_path_pathtype_leaf_directory_false
# Test-Path with -PathType Leaf returns $false for a directory
$res = Test-Path "tests" -PathType Leaf

if ($res -ne $false) {
    Write-Host "FAIL: expected `$false for directory tested with -PathType Leaf, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
