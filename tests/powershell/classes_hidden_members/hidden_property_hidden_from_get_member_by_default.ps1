# vybe-test: powershell/classes_hidden_members/hidden_property_hidden_from_get_member_by_default
class Vault {
    hidden [string]$Pin = "1234"
    [string]$User = "Owner"
}
$v = [Vault]::new()
$members = @($v | Get-Member | ForEach-Object { $_.Name })
if ($members -contains "Pin" -or -not ($members -contains "User")) {
    Write-Host "FAIL: Hidden property should not appear in default Get-Member"
    exit 1
}
Write-Host "PASS"
exit 0
