# vybe-test: powershell/parameters_validate_count/validatecount_combined_with_validateset
function Set-RestrictedChoices {
    param(
        [ValidateCount(1, 2)]
        [ValidateSet("Dev", "Test", "Prod")]
        [string[]]$Envs
    )
    return $Envs -join ":"
}
$res = Set-RestrictedChoices -Envs "Dev", "Prod"
if ($res -ne "Dev:Prod") {
    Write-Host "FAIL: ValidateCount combined with ValidateSet failed"
    exit 1
}
Write-Host "PASS"
exit 0
