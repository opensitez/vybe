# vybe-test: powershell/parameters_validate_pattern/validatepattern_whitespace_rejection
function Set-NoSpace {
    param([ValidatePattern('^\S+$')][string]$Word)
    return $Word
}
$caught = $false
try {
    $x = Set-NoSpace -Word "has space"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception on string containing whitespace"
    exit 1
}
Write-Host "PASS"
exit 0
