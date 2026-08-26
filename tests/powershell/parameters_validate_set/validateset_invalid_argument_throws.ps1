# vybe-test: powershell/parameters_validate_set/validateset_invalid_argument_throws
function Set-Env2 {
    param([ValidateSet("Dev", "Test", "Prod")][string]$EnvName)
    return $EnvName
}
$caught = $false
try {
    $x = Set-Env2 -EnvName "Staging"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected error when argument not in ValidateSet"
    exit 1
}
Write-Host "PASS"
exit 0
