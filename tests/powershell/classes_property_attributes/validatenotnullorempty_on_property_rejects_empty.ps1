# vybe-test: powershell/classes_property_attributes/validatenotnullorempty_on_property_rejects_empty
class UserAccount2 {
    [ValidateNotNullOrEmpty()][string]$Username
}
$u = [UserAccount2]::new()
$caught = $false
try {
    $u.Username = ""
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception on empty string assignment"
    exit 1
}
Write-Host "PASS"
exit 0
