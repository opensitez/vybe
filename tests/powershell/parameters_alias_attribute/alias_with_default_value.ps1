# vybe-test: powershell/parameters_alias_attribute/alias_with_default_value
function Get-ModeWithAlias {
    param([Alias("M")][string]$Mode = "Fast")
    return $Mode
}
$r1 = Get-ModeWithAlias
$r2 = Get-ModeWithAlias -M "Slow"
if ($r1 -ne "Fast" -or $r2 -ne "Slow") {
    Write-Host "FAIL: Alias with default value failed"
    exit 1
}
Write-Host "PASS"
exit 0
