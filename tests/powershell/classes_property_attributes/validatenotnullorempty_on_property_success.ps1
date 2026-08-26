# vybe-test: powershell/classes_property_attributes/validatenotnullorempty_on_property_success
class UserAccount {
    [ValidateNotNullOrEmpty()][string]$Username
}
$u = [UserAccount]::new()
$u.Username = "admin"
if ($u.Username -ne "admin") {
    Write-Host "FAIL: ValidateNotNullOrEmpty valid assignment failed"
    exit 1
}
Write-Host "PASS"
exit 0
