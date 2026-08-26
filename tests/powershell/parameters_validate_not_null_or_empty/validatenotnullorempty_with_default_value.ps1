# vybe-test: powershell/parameters_validate_not_null_or_empty/validatenotnullorempty_with_default_value
function Get-EnvName {
    param([ValidateNotNullOrEmpty()][string]$Env = "Production")
    return $Env
}
$res = Get-EnvName
if ($res -ne "Production") {
    Write-Host "FAIL: Default value with ValidateNotNullOrEmpty failed"
    exit 1
}
Write-Host "PASS"
exit 0
