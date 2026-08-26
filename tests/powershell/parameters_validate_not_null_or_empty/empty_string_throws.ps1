# vybe-test: powershell/parameters_validate_not_null_or_empty/empty_string_throws
function Set-TargetName2 {
    param([ValidateNotNullOrEmpty()][string]$Name)
    return $Name
}
$caught = $false
try {
    $x = Set-TargetName2 -Name ""
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when empty string passed to ValidateNotNullOrEmpty"
    exit 1
}
Write-Host "PASS"
exit 0
