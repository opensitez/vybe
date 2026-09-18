# vybe-test: powershell/split_path_cmdlet/split_path_isabsolute_relative_path_false
# Relative paths return $false when tested with -IsAbsolute
$res = Split-Path "relative/sub/path" -IsAbsolute

if ($res -ne $false) {
    Write-Host "FAIL: expected `$false for relative path, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
