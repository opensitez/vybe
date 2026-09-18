# vybe-test: powershell/resolve_path_cmdlet/resolve_path_relative_parameter
# Resolve-Path with -Relative returns a relative path string (e.g. .\tests or ./tests)
$rel = Resolve-Path "tests" -Relative

if ($null -eq $rel) {
    Write-Host "FAIL: Resolve-Path -Relative returned `$null"
    exit 1
}

if ($rel -notmatch "^\.[/|\\]tests$") {
    Write-Host "FAIL: expected relative path format (./tests or .\tests), got: '$rel'"
    exit 1
}

Write-Host "PASS"
exit 0
