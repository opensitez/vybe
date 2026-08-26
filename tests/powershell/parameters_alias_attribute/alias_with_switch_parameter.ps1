# vybe-test: powershell/parameters_alias_attribute/alias_with_switch_parameter
function Invoke-Build {
    param(
        [Alias("f")]
        [switch]$Force
    )
    return "Force:$($Force.IsPresent)"
}
$r1 = Invoke-Build -f
$r2 = Invoke-Build
if ($r1 -ne "Force:True" -or $r2 -ne "Force:False") {
    Write-Host "FAIL: Alias on switch parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
