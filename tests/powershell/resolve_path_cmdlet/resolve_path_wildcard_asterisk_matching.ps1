# vybe-test: powershell/resolve_path_cmdlet/resolve_path_wildcard_asterisk_matching
# Wildcards matching directories are resolved into PathInfo objects
$info = Resolve-Path "tests/pow*"

if ($null -eq $info) {
    Write-Host "FAIL: Resolve-Path with asterisk returned `$null"
    exit 1
}

if ($info.Path -notmatch "powershell$") {
    Write-Host "FAIL: expected resolved path to match 'powershell', got: '$($info.Path)'"
    exit 1
}

Write-Host "PASS"
exit 0
