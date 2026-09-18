# vybe-test: powershell/test_path_cmdlet/test_path_literal_path_with_brackets
# -LiteralPath verifies existence of files or folders whose names contain bracket characters '[' and ']'
$bracketDir = "tests/_temp_tp_[bracket_test]_dir"

[System.IO.Directory]::CreateDirectory($bracketDir) | Out-Null

$res = $false
try {
    $res = Test-Path -LiteralPath $bracketDir
} finally {
    if ([System.IO.Directory]::Exists($bracketDir)) {
        [System.IO.Directory]::Delete($bracketDir)
    }
}

if ($res -ne $true) {
    Write-Host "FAIL: Test-Path -LiteralPath returned `$false for directory containing brackets"
    exit 1
}

Write-Host "PASS"
exit 0
