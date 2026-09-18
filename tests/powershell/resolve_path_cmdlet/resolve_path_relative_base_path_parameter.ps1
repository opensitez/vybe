# vybe-test: powershell/resolve_path_cmdlet/resolve_path_relative_base_path_parameter
# -RelativeBasePath establishes the root directory against which -Relative path is computed
$rel = Resolve-Path -Path "powershell" -RelativeBasePath "tests" -Relative

if ($null -eq $rel) {
    Write-Host "FAIL: Resolve-Path with -RelativeBasePath returned `$null"
    exit 1
}

if ($rel -notmatch "^\.[/|\\]powershell$") {
    Write-Host "FAIL: expected relative path './powershell', got: '$rel'"
    exit 1
}

Write-Host "PASS"
exit 0
