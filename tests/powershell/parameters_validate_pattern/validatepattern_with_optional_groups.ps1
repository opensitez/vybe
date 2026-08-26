# vybe-test: powershell/parameters_validate_pattern/validatepattern_with_optional_groups
function Set-VersionStr {
    param([ValidatePattern('^\d+\.\d+(\.\d+)?$')][string]$V)
    return $V
}
$r1 = Set-VersionStr -V "1.0"
$r2 = Set-VersionStr -V "1.0.5"
if ($r1 -ne "1.0" -or $r2 -ne "1.0.5") {
    Write-Host "FAIL: ValidatePattern optional group failed"
    exit 1
}
Write-Host "PASS"
exit 0
