# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_in_where_object_filter
class FilterAdapter {
    static [bool]GetIsAdmin([psobject]$i) { return ($i.Role -eq "Admin") }
}
$users = @(
    [pscustomobject]@{ Role = "Admin"; Name = "Alice" },
    [pscustomobject]@{ Role = "User"; Name = "Bob" }
)
$m = [FilterAdapter].GetMethod("GetIsAdmin")
foreach ($u in $users) {
    $u.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("IsAdmin", $m))
}
$admins = @($users | Where-Object { $_.IsAdmin })
if ($admins.Length -ne 1 -or $admins[0].Name -ne "Alice") {
    Write-Host "FAIL: PSCodeProperty in Where-Object filter failed"
    exit 1
}
Write-Host "PASS"
exit 0
