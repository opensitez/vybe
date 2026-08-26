# vybe-test: powershell/parameters_validate_not_null_or_empty/validatenotnullorempty_splatting
function Set-ApiKeySplat {
    param([ValidateNotNullOrEmpty()][string]$ApiKey)
    return "OK"
}
$p = @{ ApiKey = "secret-token" }
$res = Set-ApiKeySplat @p
if ($res -ne "OK") {
    Write-Host "FAIL: Splatting with ValidateNotNullOrEmpty failed"
    exit 1
}
Write-Host "PASS"
exit 0
