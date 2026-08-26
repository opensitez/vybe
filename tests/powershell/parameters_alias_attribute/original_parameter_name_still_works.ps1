# vybe-test: powershell/parameters_alias_attribute/original_parameter_name_still_works
function Get-PortConfig {
    param(
        [Alias("P")]
        [int]$Port
    )
    return $Port
}
$r1 = Get-PortConfig -Port 8080
$r2 = Get-PortConfig -P 8080
if ($r1 -ne 8080 -or $r2 -ne 8080) {
    Write-Host "FAIL: Original parameter and alias should both bind"
    exit 1
}
Write-Host "PASS"
exit 0
