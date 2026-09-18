# vybe-test: powershell/split_path_cmdlet/split_path_trailing_separator_leaf
# When path ends with a trailing separator, -Leaf correctly returns the final named directory
$res = Split-Path "/opt/myfolder/" -Leaf

if ($res -ne "myfolder") {
    Write-Host "FAIL: expected 'myfolder', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
