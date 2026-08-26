# vybe-test: powershell/type_version_parsing_and_comparison/tryparse_valid_and_invalid
$v = $null
$ok = [version]::TryParse("5.1.0.0", [ref]$v)
$bad = [version]::TryParse("version-five", [ref]$v)
if (-not $ok -or $bad) {
    Write-Host "FAIL: TryParse check failed"
    exit 1
}
Write-Host "PASS"
exit 0
