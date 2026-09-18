# vybe-test: powershell/convert_path_cmdlet/convert_path_wildcard_asterisk_matching
# Convert-Path resolves wildcards matching files or directories
$res = Convert-Path "tests/pow*"

if ($null -eq $res) {
    Write-Host "FAIL: Convert-Path wildcard returned `$null"
    exit 1
}

if ($res -notmatch "powershell$") {
    Write-Host "FAIL: expected wildcard resolution to match 'powershell', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
