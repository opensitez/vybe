# vybe-test: powershell/null_coalescing_assignment/null_assignment_in_function
function Get-Config([string]$override) {
    $cfg = $override
    $cfg ??= "SystemDefault"
    return $cfg
}
$res1 = Get-Config $null
$res2 = Get-Config "UserOverride"
if ($res1 -ne "SystemDefault" -or $res2 -ne "UserOverride") {
    Write-Host "FAIL: function ??= expected SystemDefault and UserOverride"
    exit 1
}
Write-Host "PASS"
exit 0
