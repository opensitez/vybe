# vybe-test: powershell/parameters_validate_not_null_or_empty/validatenotnullorempty_multiple_parameters
function Set-UserGroup {
    param(
        [ValidateNotNullOrEmpty()][string]$User,
        [ValidateNotNullOrEmpty()][string]$Group
    )
    return "$User@$Group"
}
$res = Set-UserGroup -User "alice" -Group "engineers"
if ($res -ne "alice@engineers") {
    Write-Host "FAIL: Multiple ValidateNotNullOrEmpty parameters failed"
    exit 1
}
Write-Host "PASS"
exit 0
