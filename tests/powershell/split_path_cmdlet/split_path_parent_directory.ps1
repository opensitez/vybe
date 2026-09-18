# vybe-test: powershell/split_path_cmdlet/split_path_parent_directory
# Split-Path -Parent returns the directory containing the item
$res = Split-Path "/dir/sub/file.txt" -Parent

if ($res -ne "/dir/sub") {
    Write-Host "FAIL: expected '/dir/sub', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
