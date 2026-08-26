# vybe-test: powershell/classes_hidden_members/hidden_property_visible_with_force_in_get_member
class Vault2 {
    hidden [string]$Pin = "9999"
}
$v = [Vault2]::new()
$members = @($v | Get-Member -Force | ForEach-Object { $_.Name })
if (-not ($members -contains "Pin")) {
    Write-Host "FAIL: Hidden property should appear with Get-Member -Force"
    exit 1
}
Write-Host "PASS"
exit 0
