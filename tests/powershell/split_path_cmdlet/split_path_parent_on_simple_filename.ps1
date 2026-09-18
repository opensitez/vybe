# vybe-test: powershell/split_path_cmdlet/split_path_parent_on_simple_filename
# When a path has no directory separators, -Parent returns an empty string
$res = Split-Path "standalone.json" -Parent

if ($null -eq $res) {
    Write-Host "FAIL: expected empty string, got `$null"
    exit 1
}

if ($res.Length -ne 0) {
    Write-Host "FAIL: expected empty string for parent of simple filename, got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
