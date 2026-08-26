# vybe-test: powershell/parameters_validate_not_null_or_empty/validatenotnullorempty_positional_binding
function Set-PosVal {
    param([Parameter(Position=0)][ValidateNotNullOrEmpty()][string]$V)
    return $V
}
$res = Set-PosVal "first"
if ($res -ne "first") {
    Write-Host "FAIL: Positional ValidateNotNullOrEmpty failed"
    exit 1
}
Write-Host "PASS"
exit 0
