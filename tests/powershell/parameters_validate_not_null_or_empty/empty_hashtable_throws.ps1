# vybe-test: powershell/parameters_validate_not_null_or_empty/empty_hashtable_throws
function Set-ConfigHt {
    param([ValidateNotNullOrEmpty()][hashtable]$Config)
    return $Config.Count
}
$caught = $false
try {
    $x = Set-ConfigHt -Config @{}
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when empty hashtable passed to ValidateNotNullOrEmpty"
    exit 1
}
Write-Host "PASS"
exit 0
