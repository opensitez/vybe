# vybe-test: powershell/parameters_validate_pattern/validatepattern_array_parameter_one_fails_throws
function Set-Tags2 {
    param([ValidatePattern('^[a-z]+$')][string[]]$Tags)
    return $Tags
}
$caught = $false
try {
    $x = Set-Tags2 -Tags "alpha", "123", "beta"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when one array element fails ValidatePattern"
    exit 1
}
Write-Host "PASS"
exit 0
