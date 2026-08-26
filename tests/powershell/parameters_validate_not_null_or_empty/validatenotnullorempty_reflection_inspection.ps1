# vybe-test: powershell/parameters_validate_not_null_or_empty/validatenotnullorempty_reflection_inspection
function Empty-Target {
    param([ValidateNotNullOrEmpty()][string]$Field)
}
$cmd = Get-Command Empty-Target
$param = $cmd.Parameters["Field"]
$attr = $param.Attributes | Where-Object { $_.GetType().Name -eq "ValidateNotNullOrEmptyAttribute" }
if ($attr -eq $null) {
    Write-Host "FAIL: ValidateNotNullOrEmpty reflection inspection failed"
    exit 1
}
Write-Host "PASS"
exit 0
