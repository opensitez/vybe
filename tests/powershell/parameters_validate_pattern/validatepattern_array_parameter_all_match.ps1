# vybe-test: powershell/parameters_validate_pattern/validatepattern_array_parameter_all_match
function Set-Tags {
    param([ValidatePattern('^[a-z]+$')][string[]]$Tags)
    return $Tags.Length
}
$res = Set-Tags -Tags "alpha", "beta", "gamma"
if ($res -ne 3) {
    Write-Host "FAIL: ValidatePattern array parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
