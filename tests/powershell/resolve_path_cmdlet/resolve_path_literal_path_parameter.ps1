# vybe-test: powershell/resolve_path_cmdlet/resolve_path_literal_path_parameter
# -LiteralPath resolves path without expanding wildcards
$info = Resolve-Path -LiteralPath "tests"

if ($null -eq $info) {
    Write-Host "FAIL: Resolve-Path -LiteralPath returned `$null"
    exit 1
}

if ($info.Path -notmatch "tests$") {
    Write-Host "FAIL: expected resolved path ending in 'tests', got: '$($info.Path)'"
    exit 1
}

Write-Host "PASS"
exit 0
