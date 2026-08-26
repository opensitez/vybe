# vybe-test: powershell/parameters_validate_pattern/validatepattern_non_matching_email_throws
function Set-UserEmail2 {
    param([ValidatePattern('^[\w\.-]+@[\w\.-]+\.\w+$')][string]$Email)
    return $Email
}
$caught = $false
try {
    $x = Set-UserEmail2 -Email "invalid_email_at_nowhere"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when pattern does not match"
    exit 1
}
Write-Host "PASS"
exit 0
