# vybe-test: powershell/convert_path_cmdlet/convert_path_literal_path_with_brackets
# Directory paths containing bracket characters '[' and ']' are resolved correctly via -LiteralPath
$bracketDir = "tests/_temp_[bracket_test]_dir"

[System.IO.Directory]::CreateDirectory($bracketDir) | Out-Null

$resolved = $null
try {
    $resolved = Convert-Path -LiteralPath $bracketDir
} finally {
    if ([System.IO.Directory]::Exists($bracketDir)) {
        [System.IO.Directory]::Delete($bracketDir)
    }
}

if ($null -eq $resolved) {
    Write-Host "FAIL: Convert-Path -LiteralPath returned `$null for directory with brackets"
    exit 1
}

if ($resolved -notmatch "\[bracket_test\]") {
    Write-Host "FAIL: resolved literal path missing brackets: '$resolved'"
    exit 1
}

Write-Host "PASS"
exit 0
