# vybe-test: powershell/classes_property_attributes/multiple_validation_attributes_on_single_property
class SecureUsername {
    [ValidateNotNullOrEmpty()]
    [ValidateLength(3, 10)]
    [ValidatePattern('^[a-z]+$')]
    [string]$Username
}
$su = [SecureUsername]::new()
$su.Username = "admin"
$caught = $false
try {
    $su.Username = "ADMIN_123" # fails pattern and case
} catch {
    $caught = $true
}
if ($su.Username -ne "admin" -or -not $caught) {
    Write-Host "FAIL: Multiple validation attributes failed"
    exit 1
}
Write-Host "PASS"
exit 0
