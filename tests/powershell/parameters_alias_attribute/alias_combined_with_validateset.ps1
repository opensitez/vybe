# vybe-test: powershell/parameters_alias_attribute/alias_combined_with_validateset
function Set-EnvironmentMode {
    param(
        [Alias("Env")]
        [ValidateSet("Dev", "Prod")]
        [string]$Environment
    )
    return $Environment
}
$res = Set-EnvironmentMode -Env "Prod"
if ($res -ne "Prod") {
    Write-Host "FAIL: Alias combined with ValidateSet failed"
    exit 1
}
Write-Host "PASS"
exit 0
