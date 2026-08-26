# vybe-test: powershell/parameters_validate_set/validateset_reflection_attribute_inspection
function Inspect-Target {
    param([ValidateSet("A", "B")][string]$ParamA)
}
$cmd = Get-Command Inspect-Target
$param = $cmd.Parameters["ParamA"]
$attr = $param.Attributes | Where-Object { $_.GetType().Name -eq "ValidateSetAttribute" }
if ($attr -eq $null -or $attr.ValidValues.Count -ne 2) {
    Write-Host "FAIL: ValidateSet reflection attribute inspection failed"
    exit 1
}
Write-Host "PASS"
exit 0
