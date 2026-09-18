# vybe-test: powershell/split_path_cmdlet/split_path_leaf_on_simple_filename
# When a path has no directory separators, -Leaf returns the filename itself
$res = Split-Path "standalone.json" -Leaf

if ($res -ne "standalone.json") {
    Write-Host "FAIL: expected 'standalone.json', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
