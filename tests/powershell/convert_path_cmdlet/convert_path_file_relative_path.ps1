# vybe-test: powershell/convert_path_cmdlet/convert_path_file_relative_path
# Convert-Path resolves a relative file path to its absolute filesystem path
$res = Convert-Path "Cargo.toml"

if ($null -eq $res) {
    Write-Host "FAIL: Convert-Path returned `$null"
    exit 1
}

if ($res -notmatch "Cargo\.toml$") {
    Write-Host "FAIL: expected path to end with 'Cargo.toml', got: '$res'"
    exit 1
}

if (-not [System.IO.File]::Exists($res)) {
    Write-Host "FAIL: converted path is not an existing file: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
