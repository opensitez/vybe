# vybe-test: powershell/parameters_validate_range/validaterange_reflection_inspection
function Range-Target {
    param([ValidateRange(5, 25)][int]$X)
}
$cmd = Get-Command Range-Target
$param = $cmd.Parameters["X"]
$attr = $param.Attributes | Where-Object { $_.GetType().Name -eq "ValidateRangeAttribute" }
if ($attr -eq $null -or $attr.MinRange -ne 5 -or $attr.MaxRange -ne 25) {
    Write-Host "FAIL: ValidateRange reflection inspection failed"
    exit 1
}
Write-Host "PASS"
exit 0
