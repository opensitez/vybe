# vybe-test: powershell/convert_path_cmdlet/convert_path_nested_relative_subpath
# Deep multi-level relative path resolution
$res = Convert-Path "tests/powershell/engine_strict_mode"

if ($null -eq $res) {
    Write-Host "FAIL: Convert-Path returned `$null"
    exit 1
}

if ($res -notmatch "engine_strict_mode$") {
    Write-Host "FAIL: deep relative path failed to resolve, got: '$res'"
    exit 1
}

if (-not [System.IO.Directory]::Exists($res)) {
    Write-Host "FAIL: resolved deep path is not an existing directory: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
