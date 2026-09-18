# vybe-test: powershell/resolve_path_cmdlet/resolve_path_literal_path_with_brackets
# Directory paths containing bracket characters '[' and ']' are resolved correctly via -LiteralPath
$bracketDir = "tests/_temp_res_[bracket_test]_dir"

[System.IO.Directory]::CreateDirectory($bracketDir) | Out-Null

$info = $null
try {
    $info = Resolve-Path -LiteralPath $bracketDir
} finally {
    if ([System.IO.Directory]::Exists($bracketDir)) {
        [System.IO.Directory]::Delete($bracketDir)
    }
}

if ($null -eq $info) {
    Write-Host "FAIL: Resolve-Path -LiteralPath returned `$null for directory with brackets"
    exit 1
}

if ($info.Path -notmatch "\[bracket_test\]") {
    Write-Host "FAIL: resolved literal path missing brackets: '$($info.Path)'"
    exit 1
}

Write-Host "PASS"
exit 0
