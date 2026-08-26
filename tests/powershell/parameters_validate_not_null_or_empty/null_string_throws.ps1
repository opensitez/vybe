# vybe-test: powershell/parameters_validate_not_null_or_empty/null_string_throws
function Set-TargetName3 {
    param([ValidateNotNullOrEmpty()][string]$Name)
    return $Name
}
$caught = $false
try {
    $x = Set-TargetName3 -Name $null
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when null passed to ValidateNotNullOrEmpty"
    exit 1
}
Write-Host "PASS"
exit 0
