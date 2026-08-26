# vybe-test: powershell/parameters_validate_pattern/validatepattern_empty_string_rejected_if_pattern_requires_chars
function Set-NonEmptyPattern {
    param([ValidatePattern('^\w+$')][string]$Text)
    return $Text
}
$caught = $false
try {
    $x = Set-NonEmptyPattern -Text ""
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception on empty string when pattern requires chars"
    exit 1
}
Write-Host "PASS"
exit 0
