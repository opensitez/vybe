# vybe-test: powershell/split_path_cmdlet/split_path_leaf_filename
# Split-Path -Leaf returns the last element of the path (the file name with extension)
$res = Split-Path "/dir/sub/file.txt" -Leaf

if ($res -ne "file.txt") {
    Write-Host "FAIL: expected 'file.txt', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
