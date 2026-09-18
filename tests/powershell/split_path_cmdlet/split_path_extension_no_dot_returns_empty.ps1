# vybe-test: powershell/split_path_cmdlet/split_path_extension_no_dot_returns_empty
# Files without an extension return an empty string when queried with -Extension
$res = Split-Path "/dir/Makefile" -Extension

if ($null -eq $res) {
    Write-Host "FAIL: Split-Path -Extension returned `$null instead of empty string"
    exit 1
}

if ($res.Length -ne 0) {
    Write-Host "FAIL: expected empty extension for 'Makefile', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
