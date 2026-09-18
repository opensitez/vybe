# vybe-test: powershell/resolve_path_cmdlet/resolve_path_file_item_resolution
# Resolve-Path resolves existing file items to their absolute PathInfo
$info = Resolve-Path "Cargo.toml"

if ($null -eq $info) {
    Write-Host "FAIL: Resolve-Path returned `$null for Cargo.toml"
    exit 1
}

if ($info.Path -notmatch "Cargo\.toml$") {
    Write-Host "FAIL: expected resolved path to end with 'Cargo.toml', got: '$($info.Path)'"
    exit 1
}

if (-not [System.IO.File]::Exists($info.Path)) {
    Write-Host "FAIL: resolved path is not an existing file: '$($info.Path)'"
    exit 1
}

Write-Host "PASS"
exit 0
