# vybe-test: powershell/psalias_properties/psalias_property_custom_class
class ServerInfo {
    [string]$HostName = "srv01.local"
}
$s = [ServerInfo]::new()
$s | Add-Member -MemberType AliasProperty -Name "Server" -Value "HostName"
if ($s.Server -ne "srv01.local") {
    Write-Host "FAIL: AliasProperty on custom class target expected 'srv01.local', got '$($s.Server)'"
    exit 1
}
Write-Host "PASS"
exit 0
