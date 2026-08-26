# vybe-test: powershell/parameters_validate_length/validatelength_reflection_attribute_inspection
function Length-Target {
    param([ValidateLength(2, 10)][string]$Text)
}
$cmd = Get-Command Length-Target
$param = $cmd.Parameters["Text"]
$attr = $param.Attributes | Where-Object { $_.GetType().Name -eq "ValidateLengthAttribute" }
if ($attr -eq $null -or $attr.MinLength -ne 2 -or $attr.MaxLength -ne 10) {
    Write-Host "FAIL: ValidateLength reflection inspection failed"
    exit 1
}
Write-Host "PASS"
exit 0
