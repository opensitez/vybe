# vybe-test: powershell/parameters_validate_not_null_or_empty/empty_array_throws
function Set-ItemArr {
    param([ValidateNotNullOrEmpty()][string[]]$Items)
    return $Items.Length
}
$caught = $false
try {
    $x = Set-ItemArr -Items @()
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when empty array passed to ValidateNotNullOrEmpty"
    exit 1
}
Write-Host "PASS"
exit 0
